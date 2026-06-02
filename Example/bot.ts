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

    sock.ev.on('connection.update', ({ connection, lastDisconnect, qr }) => {
        if (qr) {
            qrcode.generate(qr, { small: true })
            console.log('Сканируй QR в WhatsApp → Связанные устройства')
        }
        if (connection === 'open') {
            logger.info({ me: sock.user?.id }, '✅ connected')
        }
        if (connection === 'close') {
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

            await pushToCrm({
                msgId: msg.key.id || '',
                from: msg.key.remoteJid,
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
