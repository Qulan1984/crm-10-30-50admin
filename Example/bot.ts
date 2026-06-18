import { Boom } from '@hapi/boom'
import qrcode from 'qrcode-terminal'
import makeWASocket, {
    DisconnectReason,
    fetchLatestBaileysVersion,
    isJidNewsletter,
    makeCacheableSignalKeyStore,
    useMultiFileAuthState
} from '../src'
import P from 'pino'

// ── CONFIG ─────────────────────────────────────────────────────────────
// Paste your Apps Script Web App URL here after deploy.
const GS_URL = process.env.GS_URL || 'https://script.google.com/macros/s/AKfycbwNCnfHq5L1RBt1rOF-MG6cjrl9hsqMKArdYC5FR_Xb6L-CENsi3s3P_B3EbcN_3bWx/exec'
// ───────────────────────────────────────────────────────────────────────

const logger = P({ transport: { target: 'pino-pretty', options: { colorize: true } } })
logger.level = 'info'

const sentToCrm = new Set<string>()

async function pushToCrm(payload: {
    msgId: string
    from: string
    pushName: string
    text: string
    timestamp: number
}) {
    if (GS_URL.startsWith('PASTE_')) {
        logger.warn('GS_URL is not set — skipping push to CRM')
        return
    }
    if (sentToCrm.has(payload.msgId)) return
    sentToCrm.add(payload.msgId)

    try {
        const res = await fetch(GS_URL + '?action=addMessage', {
            method: 'POST',
            body: JSON.stringify(payload),
            redirect: 'follow'
        })
        const json: any = await res.json().catch(() => ({}))
        logger.info({ from: payload.from, isNewLead: json?.isNewLead }, '→ CRM')
    } catch (err) {
        sentToCrm.delete(payload.msgId)
        logger.error({ err }, 'failed to push to CRM')
    }
}

// ── OUTGOING: drain the CRM outbox and send messages back to WhatsApp ──
const OUTBOX_POLL_MS = 3000
const sentOutbox = new Set<string>()   // ids already handled this process — dedup
let draining = false

async function markSent(id: string, status: 'sent' | 'failed') {
    try {
        await fetch(GS_URL + '?action=markSent', {
            method: 'POST',
            body: JSON.stringify({ id, status }),
            redirect: 'follow'
        })
    } catch (err) {
        logger.error({ err, id }, 'markSent failed')
    }
}

async function drainOutbox(sock: any) {
    if (GS_URL.startsWith('PASTE_') || draining) return
    draining = true
    try {
        const res = await fetch(GS_URL + '?action=pollOutbox', { redirect: 'follow' })
        const json: any = await res.json().catch(() => ({}))
        const items: Array<{ id: string; phone: string; text: string }> = json?.outbox || []
        for (const it of items) {
            if (!it.id || sentOutbox.has(it.id)) continue
            sentOutbox.add(it.id)
            const digits = String(it.phone || '').replace(/\D/g, '')
            if (!digits || !it.text) { await markSent(it.id, 'failed'); continue }
            try {
                await sock.sendMessage(digits + '@s.whatsapp.net', { text: it.text })
                await markSent(it.id, 'sent')
                logger.info({ to: digits }, '→ WhatsApp (from CRM)')
            } catch (err) {
                sentOutbox.delete(it.id)
                logger.error({ err, id: it.id }, 'failed to send outgoing message')
                await markSent(it.id, 'failed')
            }
        }
    } catch (err) {
        logger.error({ err }, 'pollOutbox failed')
    } finally {
        draining = false
    }
}

// WhatsApp now addresses many chats by LID (<id>@lid) instead of the phone number.
// Resolve it to the real phone (PN) so the CRM keys contacts by phone, not LID —
// otherwise calls (which come with the real number) never match WhatsApp leads.
async function resolveSenderPhone(sock: any, key: any): Promise<string> {
    const jid: string = key.remoteJid || ''
    if (jid.endsWith('@s.whatsapp.net')) return jid
    if (jid.endsWith('@lid')) {
        if (typeof key.remoteJidAlt === 'string' && key.remoteJidAlt.endsWith('@s.whatsapp.net')) {
            return key.remoteJidAlt
        }
        try {
            const pn = await sock.signalRepository?.lidMapping?.getPNForLID(jid)
            if (pn) return pn
        } catch {
            /* mapping not known yet — fall back to the LID */
        }
    }
    return jid
}

function extractText(message: any): string {
    if (!message) return ''
    return (
        message.conversation ||
        message.extendedTextMessage?.text ||
        message.imageMessage?.caption ||
        message.videoMessage?.caption ||
        ''
    )
}

const start = async () => {
    const { state, saveCreds } = await useMultiFileAuthState('baileys_auth_info')
    const { version } = await fetchLatestBaileysVersion()

    const sock = makeWASocket({
        version,
        logger,
        qrTimeout: 90_000,
        auth: {
            creds: state.creds,
            keys: makeCacheableSignalKeyStore(state.keys, logger)
        }
    })

    sock.ev.on('creds.update', saveCreds)

    let outboxTimer: ReturnType<typeof setInterval> | null = null

    sock.ev.on('connection.update', ({ connection, lastDisconnect, qr }) => {
        if (qr) {
            qrcode.generate(qr, { small: true })
            console.log('Сканируй QR в WhatsApp → Связанные устройства')
        }
        if (connection === 'open') {
            logger.info({ me: sock.user?.id }, '✅ connected')
            // start draining the CRM outbox (clear any prior timer from a previous socket)
            if (outboxTimer) clearInterval(outboxTimer)
            outboxTimer = setInterval(() => drainOutbox(sock), OUTBOX_POLL_MS)
        }
        if (connection === 'close') {
            if (outboxTimer) { clearInterval(outboxTimer); outboxTimer = null }
            const code = (lastDisconnect?.error as Boom)?.output?.statusCode
            const shouldReconnect = code !== DisconnectReason.loggedOut
            logger.warn({ code, shouldReconnect }, 'connection closed')
            if (shouldReconnect) start()
        }
    })

    sock.ev.on('messages.upsert', async ({ messages, type }) => {
        if (type !== 'notify') return

        for (const msg of messages) {
            if (msg.key.fromMe) continue
            if (!msg.key.remoteJid) continue
            if (isJidNewsletter(msg.key.remoteJid)) continue
            // skip groups for now — uncomment to enable
            if (msg.key.remoteJid.endsWith('@g.us')) continue

            const text = extractText(msg.message)
            if (!text) continue

            const from = await resolveSenderPhone(sock, msg.key)

            await pushToCrm({
                msgId: msg.key.id || '',
                from,
                pushName: msg.pushName || '',
                text,
                timestamp: Number(msg.messageTimestamp) * 1000 || Date.now()
            })
        }
    })
}

start().catch(err => {
    logger.error({ err }, 'fatal')
    process.exit(1)
})
