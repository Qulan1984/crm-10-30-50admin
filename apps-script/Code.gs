/**
 * CRM 10-30-50 backend on Google Apps Script.
 *
 * Sheets expected in the bound spreadsheet:
 *   - "funnel"               : id | name | phone | desc | status | touches | comments | tasks
 *   - "Потенциальные" etc.   : id | name | phone | desc | touches | comments | tasks
 *   - "Messages" (auto)      : id | timestamp | from | pushName | text | isNewLead | dir | status
 *                              (dir = in|out ; status = ''|pending|sent|failed ; one row per chat message)
 *
 * Endpoints:
 *   GET  ?action=readFunnel
 *   GET  ?action=readOther
 *   GET  ?action=pollMessages&since=<unix_ms>     -> incoming messages only (for toasts/lead creation)
 *   GET  ?action=thread&phone=<phone>             -> full chat thread with one contact (in + out)
 *   GET  ?action=pollOutbox                        -> pending outgoing messages (bot drains this)
 *   GET  ?action=pollCalls&since=<unix_ms>
 *   GET  ?action=callRecording&id=<recordId>     -> {mime, b64} (proxied from Beeline)
 *   POST ?action=writeAll       body: {funnel: rows[][], other: {sheetName: rows[][]}}
 *   POST ?action=addMessage     body: {from, text, pushName, timestamp, msgId}
 *   POST ?action=sendMessage    body: {phone, text}                 -> queues an outgoing message
 *   POST ?action=markSent       body: {id, status}                  -> bot marks an outgoing msg sent/failed
 *
 * Beeline Cloud PBX (KZ) call sync: see the block at the bottom of this file.
 */

const MESSAGES_SHEET = 'Messages';
const FUNNEL_SHEET = 'funnel';
const CALLS_SHEET = 'Calls';
const KNOWN_OTHER = ['Потенциальные', 'Сделка', 'КХ'];

// updatedAt is the poll cursor (bumped on every change); timestamp stays the real call start.
const CALLS_HEADER = ['callId', 'timestamp', 'direction', 'phone', 'abonent', 'status', 'duration', 'recordId', 'comment', 'updatedAt'];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'readFunnel')    return json({ funnel: readSheet_(FUNNEL_SHEET) });
    if (action === 'readOther')     return json({ other: readOther_() });
    if (action === 'pollMessages')  return json({ messages: pollMessages_(Number(e.parameter.since || 0)) });
    if (action === 'thread')        return json({ thread: thread_(e.parameter.phone || '') });
    if (action === 'pollOutbox')    return json({ outbox: pollOutbox_() });
    if (action === 'pollCalls')     return json({ calls: pollCalls_(Number(e.parameter.since || 0)) });
    if (action === 'callRecording') return json(callRecording_(e.parameter.id));
    return json({ error: 'unknown action', action });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
  try {
    if (action === 'writeAll')     { writeAll_(body); return json({ ok: true }); }
    if (action === 'addMessage')   return json(addMessage_(body));
    if (action === 'sendMessage')  return json(sendMessage_(body));
    if (action === 'markSent')     return json(markSent_(body));
    return json({ error: 'unknown action', action });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return [];
  return sh.getDataRange().getValues();
}

function readOther_() {
  const ss = SpreadsheetApp.getActive();
  const out = {};
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === FUNNEL_SHEET || name === MESSAGES_SHEET) return;
    out[name] = sh.getDataRange().getValues();
  });
  return out;
}

function writeAll_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();

    if (body.funnel && body.funnel.length) {
      const sh = getOrCreate_(FUNNEL_SHEET);
      sh.clear();
      sh.getRange(1, 1, body.funnel.length, body.funnel[0].length).setValues(body.funnel);
    }

    if (body.other) {
      Object.keys(body.other).forEach(name => {
        const rows = body.other[name];
        if (!rows || !rows.length) return;
        const sh = getOrCreate_(name);
        sh.clear();
        sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      });
    }
  } finally {
    lock.releaseLock();
  }
}

function getOrCreate_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Bot calls this on every incoming WhatsApp message. */
function addMessage_(m) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getOrCreate_(MESSAGES_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['id', 'timestamp', 'from', 'pushName', 'text', 'isNewLead', 'dir', 'status']);
    }

    const phone = normalizePhone_(m.from || '');
    const isNewLead = phone && !phoneExistsAnywhere_(phone);

    sh.appendRow([
      m.msgId || '',
      Number(m.timestamp) || Date.now(),
      phone,
      m.pushName || '',
      m.text || '',
      isNewLead ? 'true' : 'false',
      'in',
      ''
    ]);

    return { ok: true, isNewLead };
  } finally {
    lock.releaseLock();
  }
}

/** CRM polls this. Returns INCOMING messages with timestamp > since (skips our own sent ones). */
function pollMessages_(since) {
  const sh = SpreadsheetApp.getActive().getSheetByName(MESSAGES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[6] || 'in') === 'out') continue; // don't surface our own sent messages as incoming
    const ts = Number(r[1]) || 0;
    if (ts > since) {
      out.push({
        id: r[0],
        timestamp: ts,
        from: r[2],
        pushName: r[3],
        text: r[4],
        isNewLead: String(r[5]) === 'true'
      });
    }
  }
  return out;
}

/** Full chat thread (incoming + outgoing) with one contact, matched by last 10 phone digits. */
function thread_(phone) {
  const sh = SpreadsheetApp.getActive().getSheetByName(MESSAGES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const key = phoneKey_(phone);
  if (!key) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (phoneKey_(r[2]) !== key) continue;
    out.push({
      id: r[0],
      timestamp: Number(r[1]) || 0,
      text: r[4],
      dir: String(r[6] || 'in'),
      status: String(r[7] || ''),
      pushName: r[3]
    });
  }
  out.sort((a, b) => a.timestamp - b.timestamp);
  return out;
}

/** CRM queues an outgoing message (dir=out, status=pending). The bot drains it via pollOutbox. */
function sendMessage_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getOrCreate_(MESSAGES_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['id', 'timestamp', 'from', 'pushName', 'text', 'isNewLead', 'dir', 'status']);
    }
    const phone = normalizePhone_(body.phone || '');
    if (!phone) return { ok: false, error: 'no phone' };
    const id = 'out_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    const ts = Date.now();
    sh.appendRow([id, ts, phone, '', body.text || '', 'false', 'out', 'pending']);
    return { ok: true, id: id, timestamp: ts };
  } finally {
    lock.releaseLock();
  }
}

/** Bot polls this to get messages it still needs to send. */
function pollOutbox_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MESSAGES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[6]) === 'out' && String(r[7]) === 'pending') {
      out.push({ id: r[0], phone: r[2], text: r[4] });
    }
  }
  return out;
}

/** Bot calls this after it sends (or fails to send) an outgoing message. */
function markSent_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(MESSAGES_SHEET);
    if (!sh || sh.getLastRow() < 2) return { ok: false };
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(body.id)) {
        sh.getRange(i + 1, 8).setValue(body.status || 'sent'); // column 8 = status
        return { ok: true };
      }
    }
    return { ok: false, error: 'not found' };
  } finally {
    lock.releaseLock();
  }
}

/** Last 10 digits of a phone — unifies 8.../+7.../77... formats (mirrors phoneKey on the frontend). */
function phoneKey_(p) {
  const d = String(p == null ? '' : p).replace(/\D/g, '');
  return d.length > 10 ? d.slice(-10) : d;
}

function normalizePhone_(jidOrPhone) {
  // strip "@s.whatsapp.net" / "@g.us" / non-digits
  const s = String(jidOrPhone).split('@')[0];
  return s.replace(/\D/g, '');
}

function phoneExistsAnywhere_(phone) {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const name = sh.getName();
    if (name === MESSAGES_SHEET) continue;
    if (sh.getLastRow() < 2) continue;
    // phone is column C (index 3) in both funnel and other sheets
    const col = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
    for (let j = 0; j < col.length; j++) {
      if (normalizePhone_(col[j][0]) === phone) return true;
    }
  }
  return false;
}

// ============================================================================
// Beeline Cloud PBX (KZ portal). There is no public REST API on cloudpbx.beeline.kz,
// so we drive the same internal endpoints the web portal itself uses, authenticating
// with the BearerToken cookie obtained by logging in. This is read-only.
//
// Script Properties (File > Project Settings > Script Properties):
//   BEELINE_LOGIN      - portal login
//   BEELINE_PASSWORD   - portal password
//   BEELINE_PROFILE_ID - profile id from the portal URLs (e.g. 2991)
//   BEELINE_BASE       - optional, default https://cloudpbx.beeline.kz/VPBX/
//
// Recurring: add a time-driven trigger on syncBeelineCalls() (every 1-5 min). It pulls
// the call journal + recordings list and upserts them into the Calls sheet; the CRM
// polls pollCalls and streams audio through callRecording (GetCallRecordContent).
// ============================================================================

const BEELINE_BASE_DEFAULT = 'https://cloudpbx.beeline.kz/VPBX/';
const BEELINE_LOOKBACK_DAYS = 3; // re-scan recent window each run so statuses/records fill in

function beelineProp_(name, def) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  return (v === null || v === '') ? def : v;
}

function beelineBase_() {
  let b = beelineProp_('BEELINE_BASE', BEELINE_BASE_DEFAULT);
  if (b.slice(-1) !== '/') b += '/';
  return b;
}

function beelineProfileId_() {
  const p = beelineProp_('BEELINE_PROFILE_ID', '');
  if (!p) throw new Error('BEELINE_PROFILE_ID script property is not set');
  return p;
}

/** Log in, capture the BearerToken cookie, cache it for reuse. */
function beelineLogin_() {
  const login = beelineProp_('BEELINE_LOGIN', '');
  const pass = beelineProp_('BEELINE_PASSWORD', '');
  if (!login || !pass) throw new Error('BEELINE_LOGIN / BEELINE_PASSWORD not set');

  const resp = UrlFetchApp.fetch(beelineBase_() + 'Account/Login', {
    method: 'post',
    followRedirects: false,
    muteHttpExceptions: true,
    payload: { Login: login, Password: pass, RememberMe: 'false' }
  });
  const cookie = extractBearerCookie_(resp.getAllHeaders());
  if (!cookie) throw new Error('Beeline login failed (' + resp.getResponseCode() + ')');
  CacheService.getScriptCache().put('BEELINE_COOKIE', cookie, 1500); // 25 min
  return cookie;
}

function beelineCookie_() {
  return CacheService.getScriptCache().get('BEELINE_COOKIE') || beelineLogin_();
}

/** The cookie value is raw JSON ({"AccessToken":...}); grab it whole. */
function extractBearerCookie_(headers) {
  let sc = headers['Set-Cookie'];
  if (!sc) return '';
  if (Array.isArray(sc)) sc = sc.join('\n');
  const m = sc.match(/BearerToken=(\{.*?\})/);
  return m ? 'BearerToken=' + m[1] : '';
}

/** Authenticated request to a portal path; re-login once on an auth redirect. */
function beelineFetch_(path, opts, retry) {
  const o = opts || {};
  o.followRedirects = false;
  o.muteHttpExceptions = true;
  o.headers = Object.assign({ Cookie: beelineCookie_() }, o.headers || {});
  const resp = UrlFetchApp.fetch(beelineBase_() + path, o);
  const code = resp.getResponseCode();
  if ((code === 302 || code === 401) && !retry) {
    CacheService.getScriptCache().remove('BEELINE_COOKIE');
    beelineLogin_();
    return beelineFetch_(path, opts, true);
  }
  return resp;
}

function beelinePostJson_(path, payload) {
  const resp = beelineFetch_(path, { method: 'post', payload: payload });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Beeline ' + path + ' -> ' + code + ' ' + resp.getContentText().slice(0, 200));
  return JSON.parse(resp.getContentText());
}

function beelineWindow_() {
  const end = new Date();
  const start = new Date(end.getTime() - BEELINE_LOOKBACK_DAYS * 86400000);
  const fmt = function (d) { return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss"); };
  return { start: fmt(start), end: fmt(end) };
}

/** "2026-06-17 14:21:59.000" (UTC, per portal) -> epoch ms. */
function beelineParseDt_(s) {
  if (!s) return Date.now();
  const d = new Date(String(s).replace(' ', 'T').replace(/\.\d+$/, '') + 'Z');
  const t = d.getTime();
  return isNaN(t) ? Date.now() : t;
}

/**
 * Pull the call journal + recordings list for the recent window and upsert into Calls.
 * Run on a time-driven trigger. Dedupes by CallID, so overlapping windows are safe.
 */
function syncBeelineCalls() {
  const pid = beelineProfileId_();
  const w = beelineWindow_();

  const calls = beelinePostJson_('Stat/CallLog?query.ProfileID=' + pid, {
    limit: '1000', page: '1', ascending: '0', orderBy: 'StartDT', byColumn: '0',
    'query.StartDT.start': w.start, 'query.StartDT.end': w.end
  });
  (calls.data || []).forEach(function (r) {
    upsertCall_({
      callId: r.CallID,
      timestamp: beelineParseDt_(r.StartDT),
      direction: r.Direction,
      phone: normalizePhone_(r.Number),
      abonent: r.AbonentName || '',
      status: r.IsMissed ? ('Пропущенный · ' + (r.StatusStr || '')) : (r.StatusStr || r.Status || ''),
      duration: r.Duration || 0
    });
  });

  const recs = beelinePostJson_('Cloud/CallRecord?query.ProfileID=' + pid, {
    limit: '1000', page: '1', ascending: '0', orderBy: 'DT', byColumn: '0',
    'query.DT.start': w.start, 'query.DT.end': w.end
  });
  (recs.data || []).forEach(function (r) {
    // recordId is the CallRecord row ID consumed by GetCallRecordContent.
    upsertCall_({ callId: r.CallID, recordId: r.ID, comment: r.Comment || '' });
  });

  return { calls: (calls.data || []).length, records: (recs.data || []).length };
}

/** Insert or update a Calls row by callId; only overwrites fields actually provided. */
function upsertCall_(c) {
  if (!c.callId) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getOrCreate_(CALLS_SHEET);
    if (sh.getLastRow() === 0) sh.appendRow(CALLS_HEADER);

    const col = { callId: 1, timestamp: 2, direction: 3, phone: 4, abonent: 5, status: 6, duration: 7, recordId: 8, comment: 9, updatedAt: 10 };
    const now = Date.now();
    let rowIndex = 0;
    if (sh.getLastRow() >= 2) {
      const ids = sh.getRange(2, col.callId, sh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(c.callId)) { rowIndex = i + 2; break; }
      }
    }

    if (!rowIndex) {
      sh.appendRow([
        c.callId, c.timestamp || now, c.direction || '', c.phone || '',
        c.abonent || '', c.status || '', c.duration || '', c.recordId || '', c.comment || '', now
      ]);
      return;
    }
    // Only bump updatedAt when a provided field actually differs, so unchanged
    // rows don't re-surface to the CRM on every sync.
    let changed = false;
    const set = function (key, val) {
      if (val === undefined || val === '' || val === null) return;
      const cell = sh.getRange(rowIndex, col[key]);
      if (String(cell.getValue()) !== String(val)) { cell.setValue(val); changed = true; }
    };
    set('timestamp', c.timestamp); set('direction', c.direction); set('phone', c.phone);
    set('abonent', c.abonent); set('status', c.status); set('duration', c.duration);
    set('recordId', c.recordId); set('comment', c.comment);
    if (changed) sh.getRange(rowIndex, col.updatedAt).setValue(now);
  } finally {
    lock.releaseLock();
  }
}

/** CRM polls this; returns calls with timestamp > since. */
function pollCalls_(since) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CALLS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const ts = Number(r[1]) || 0;
    const upd = Number(r[9]) || ts; // poll cursor
    if (upd > since) {
      out.push({
        callId: r[0], timestamp: ts, direction: r[2], phone: r[3],
        abonent: r[4], status: r[5], duration: r[6], recordId: r[7], comment: r[8], updatedAt: upd
      });
    }
  }
  return out;
}

/** Proxy recording audio so the browser never holds portal credentials. */
function callRecording_(recordId) {
  if (!recordId) return { error: 'no record id' };
  const resp = beelineFetch_('Cloud/GetCallRecordContent?asPreview=false&id=' + encodeURIComponent(recordId));
  if (resp.getResponseCode() !== 200) return { error: 'beeline ' + resp.getResponseCode() };
  const blob = resp.getBlob();
  return { mime: 'audio/mpeg', b64: Utilities.base64Encode(blob.getBytes()) };
}

/** Quick manual check from the editor: logs in and returns the recent call count. */
function testBeelineLogin() {
  beelineLogin_();
  const r = syncBeelineCalls();
  Logger.log(JSON.stringify(r));
  return r;
}

/**
 * ONE-CLICK SETUP. Fill the three values below, then press Run on this function once.
 * It saves the credentials to Script Properties, creates the every-minute sync trigger,
 * and does a first sync. After it works, blank the values back out so you don't commit them.
 */
function setupBeelineTelephony() {
  const LOGIN = '';       // <- логин от портала
  const PASSWORD = '';    // <- пароль от портала
  const PROFILE_ID = '';  // <- например 2991

  if (!LOGIN || !PASSWORD || !PROFILE_ID) {
    throw new Error('Заполни LOGIN, PASSWORD и PROFILE_ID в начале функции setupBeelineTelephony, потом нажми Выполнить.');
  }

  PropertiesService.getScriptProperties().setProperties({
    BEELINE_LOGIN: LOGIN,
    BEELINE_PASSWORD: PASSWORD,
    BEELINE_PROFILE_ID: String(PROFILE_ID)
  });

  // Recreate the sync trigger (remove duplicates first).
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncBeelineCalls') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncBeelineCalls').timeBased().everyMinutes(1).create();

  const r = syncBeelineCalls();
  Logger.log('Готово. Синхронизировано: ' + JSON.stringify(r) + '. Триггер на syncBeelineCalls создан.');
  return r;
}
