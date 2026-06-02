/**
 * CRM 10-30-50 backend on Google Apps Script.
 *
 * Sheets expected in the bound spreadsheet:
 *   - "funnel"               : id | name | phone | desc | status | touches | comments | tasks
 *   - "Потенциальные" etc.   : id | name | phone | desc | touches | comments | tasks
 *   - "Messages" (auto)      : id | timestamp | from | pushName | text | isNewLead
 *
 * Endpoints:
 *   GET  ?action=readFunnel
 *   GET  ?action=readOther
 *   GET  ?action=pollMessages&since=<unix_ms>
 *   POST ?action=writeAll       body: {funnel: rows[][], other: {sheetName: rows[][]}}
 *   POST ?action=addMessage     body: {from, text, pushName, timestamp, msgId}
 */

const MESSAGES_SHEET = 'Messages';
const FUNNEL_SHEET = 'funnel';
const KNOWN_OTHER = ['Потенциальные', 'Сделка', 'КХ'];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'readFunnel')   return json({ funnel: readSheet_(FUNNEL_SHEET) });
    if (action === 'readOther')    return json({ other: readOther_() });
    if (action === 'pollMessages') return json({ messages: pollMessages_(Number(e.parameter.since || 0)) });
    return json({ error: 'unknown action', action });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
  try {
    if (action === 'writeAll')   { writeAll_(body); return json({ ok: true }); }
    if (action === 'addMessage') return json(addMessage_(body));
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
      sh.appendRow(['id', 'timestamp', 'from', 'pushName', 'text', 'isNewLead']);
    }

    const phone = normalizePhone_(m.from || '');
    const isNewLead = phone && !phoneExistsAnywhere_(phone);

    sh.appendRow([
      m.msgId || '',
      Number(m.timestamp) || Date.now(),
      phone,
      m.pushName || '',
      m.text || '',
      isNewLead ? 'true' : 'false'
    ]);

    return { ok: true, isNewLead };
  } finally {
    lock.releaseLock();
  }
}

/** CRM polls this. Returns messages with timestamp > since. */
function pollMessages_(since) {
  const sh = SpreadsheetApp.getActive().getSheetByName(MESSAGES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
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
