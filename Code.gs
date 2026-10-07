/**
 * Class check-in backend (Google Apps Script)
 * 1) In your Google Sheet: Extensions > Apps Script. Paste this file and save.
 * 2) Run setup() once, allow access, and copy the teacher key from the execution log.
 * 3) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 */
const TZ = 'Asia/Dubai';
const WINDOW_SEC = 30;    // QR code changes every 30 seconds
const VALID_WINDOWS = 3;  // a scanned code is accepted for about 60-90 seconds
const WRITE_ACTIONS = { approve: 1, device: 1, scan: 1, register: 1 };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, 'Roster', ['Class', 'Student']);
  ensureSheet_(ss, 'Checkins', ['Date', 'Time', 'Class', 'Student', 'Device', 'Late']).getRange('A:F').setNumberFormat('@');
  ensureSheet_(ss, 'Devices', ['Device', 'Student', 'Class', 'Status', 'Requested', 'Note']).getRange('A:F').setNumberFormat('@');
  const log = ensureSheet_(ss, 'Daily Log', ['Date', 'Class', 'Present', 'Absent', 'Absentees', 'Late arrivals', 'Saved at']);
  log.getRange('A:B').setNumberFormat('@');
  log.getRange('E:G').setNumberFormat('@');
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('SECRET', secret);
  }
  Logger.log('Your teacher key: ' + secret);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

/* ---------- web endpoints ---------- */
function doGet() { return json_({ ok: true, msg: 'Attendance backend is running' }); }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (WRITE_ACTIONS[req.action]) {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { out = handle_(req); } finally { lock.releaseLock(); }
    } else {
      out = handle_(req);
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  switch (req.action) {
    // teacher
    case 'classes': requireTeacher_(req); return { ok: true, classes: classes_() };
    case 'status':  requireTeacher_(req); return status_(String(req.cls));
    case 'approve': requireTeacher_(req); return approve_(req);
    case 'device':  requireTeacher_(req); return decideDevice_(req);
    // student
    case 'scan':     return scan_(req);
    case 'register': return register_(req);
    case 'mystate':  return myState_(req);
    default: throw new Error('Unknown action');
  }
}

/* ---------- helpers ---------- */
function secret_() {
  const s = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!s) throw new Error('Run setup() in Apps Script first.');
  return s;
}
function requireTeacher_(req) { if (String(req.secret) !== secret_()) throw new Error('Wrong teacher key.'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowTime_() { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function windowNow_() { return Math.floor(Date.now() / 1000 / WINDOW_SEC); }

function codeFor_(cls, w) {
  const b = Utilities.computeHmacSha256Signature(cls + '|' + w, secret_());
  const n = (b[0] & 0xff) * 16777216 + (b[1] & 0xff) * 65536 + (b[2] & 0xff) * 256 + (b[3] & 0xff);
  return ('000000' + (n % 1000000)).slice(-6);
}

function requireCode_(cls, code) {
  const w = windowNow_();
  for (let i = 0; i < VALID_WINDOWS; i++) if (codeFor_(cls, w - i) === String(code)) return;
  throw new Error('This code has expired. Scan the code on the board again.');
}

function table_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing "' + name + '" tab. Run setup() first.');
  return { sh: sh, rows: sh.getDataRange().getDisplayValues().slice(1) };
}

function classes_() {
  const set = {};
  table_('Roster').rows.forEach(r => { const c = r[0].trim(); if (c && r[1].trim()) set[c] = true; });
  return Object.keys(set).sort();
}

function roster_(cls) {
  return table_('Roster').rows
    .filter(r => r[0].trim() === cls && r[1].trim())
    .map(r => r[1].trim())
    .sort((a, b) => a.localeCompare(b));
}

function checkins_(cls, date) {
  return table_('Checkins').rows
    .filter(r => r[0] === date && r[2] === cls)
    .map(r => ({ time: r[1], name: r[3], late: r[5] === 'yes' }));
}

function findLog_(t, date, cls) { return t.rows.findIndex(r => r[0] === date && r[1] === cls); }

function splitNames_(s) { return s ? s.split(', ').filter(Boolean) : []; }

/** Records a check-in once per student per day. After the register is saved, it counts as late. */
function record_(cls, name, device) {
  const date = today_();
  const existing = checkins_(cls, date).find(x => x.name === name);
  if (existing) return { time: existing.time, late: existing.late, already: true };
  const log = table_('Daily Log');
  const i = findLog_(log, date, cls);
  const time = nowTime_();
  const late = i >= 0;
  table_('Checkins').sh.appendRow([date, time, cls, name, device || 'teacher', late ? 'yes' : '']);
  if (late) {
    const r = log.rows[i];
    const absentees = splitNames_(r[4]);
    const wasAbsent = absentees.indexOf(name) >= 0;
    const remaining = absentees.filter(n => n !== name);
    const lateList = splitNames_(r[5]).concat([name + ' (' + time + ')']);
    log.sh.getRange(i + 2, 3, 1, 4).setValues([[
      Number(r[2]) + (wasAbsent ? 1 : 0), remaining.length, remaining.join(', '), lateList.join(', ')
    ]]);
  }
  return { time: time, late: late };
}

/* ---------- devices ---------- */
function findDevice_(device) {
  const t = table_('Devices');
  const i = t.rows.findIndex(r => r[0] === device);
  return i < 0 ? null : { t: t, i: i, device: t.rows[i][0], name: t.rows[i][1], cls: t.rows[i][2], status: t.rows[i][3] };
}

function deviceReply_(d, cls, doCheckin) {
  if (d.status === 'pending') return { ok: true, state: 'pending', name: d.name };
  if (d.status === 'rejected') return { ok: true, state: 'rejected' };
  if (d.status === 'replaced') return { ok: true, state: 'replaced', name: d.name };
  if (d.cls !== cls) throw new Error('This tablet is linked to class ' + d.cls + '.');
  if (doCheckin) {
    const r = record_(cls, d.name, d.device);
    return { ok: true, state: 'in', name: d.name, time: r.time, late: r.late };
  }
  const today = checkins_(cls, today_()).find(x => x.name === d.name);
  return today
    ? { ok: true, state: 'in', name: d.name, time: today.time, late: today.late }
    : { ok: true, state: 'approved', name: d.name };
}

/* ---------- student actions ---------- */
function scan_(req) {
  const cls = String(req.cls), device = String(req.device || '');
  requireCode_(cls, req.code);
  if (!device) throw new Error('This browser cannot save data. Turn off private browsing and scan again.');
  const d = findDevice_(device);
  if (!d) return { ok: true, state: 'new', roster: roster_(cls) };
  return deviceReply_(d, cls, true);
}

function register_(req) {
  const cls = String(req.cls), device = String(req.device || ''), name = String(req.name || '');
  requireCode_(cls, req.code);
  if (!device) throw new Error('This browser cannot save data. Turn off private browsing and scan again.');
  const d = findDevice_(device);
  if (d) return deviceReply_(d, cls, false);
  if (roster_(cls).indexOf(name) === -1) throw new Error('Choose your name from the list.');
  table_('Devices').sh.appendRow([device, name, cls, 'pending', today_() + ' ' + nowTime_(), '']);
  return { ok: true, state: 'pending', name: name };
}

function myState_(req) {
  const d = findDevice_(String(req.device || ''));
  if (!d) return { ok: true, state: 'none' };
  return deviceReply_(d, String(req.cls), false);
}

/* ---------- teacher actions ---------- */
function status_(cls) {
  const date = today_();
  const nowSec = Math.floor(Date.now() / 1000);
  const devices = table_('Devices').rows.filter(r => r[2] === cls);
  const linked = devices.filter(r => r[3] === 'approved').map(r => r[1]);
  const pending = devices.filter(r => r[3] === 'pending').map(r => ({
    device: r[0], name: r[1], requested: r[4], second: linked.indexOf(r[1]) >= 0
  }));
  const log = table_('Daily Log');
  const i = findLog_(log, date, cls);
  return {
    ok: true, cls: cls, date: date,
    code: codeFor_(cls, windowNow_()),
    secondsLeft: WINDOW_SEC - (nowSec % WINDOW_SEC),
    roster: roster_(cls),
    checkedIn: checkins_(cls, date),
    linked: linked,
    pending: pending,
    savedAt: i >= 0 ? log.rows[i][6] : ''
  };
}

function decideDevice_(req) {
  const d = findDevice_(String(req.device));
  if (!d) throw new Error('That request no longer exists.');
  if (req.decision === 'approve') {
    d.t.rows.forEach((r, j) => {
      if (j !== d.i && r[1] === d.name && r[2] === d.cls && r[3] === 'approved') {
        d.t.sh.getRange(j + 2, 4, 1, 3).setValues([['replaced', r[4], 'Replaced ' + today_()]]);
      }
    });
    d.t.sh.getRange(d.i + 2, 4).setValue('approved');
    // Approving during today's lesson also checks the student in.
    if (String(d.t.rows[d.i][4]).indexOf(today_()) === 0) record_(d.cls, d.name, d.device);
  } else {
    d.t.sh.getRange(d.i + 2, 4).setValue('rejected');
  }
  return { ok: true };
}

function approve_(req) {
  const cls = String(req.cls);
  const absent = (req.absent || []).map(String);
  const present = Number(req.present) || 0;
  const date = today_();
  const log = table_('Daily Log');
  const i = findLog_(log, date, cls);
  const late = i >= 0 ? log.rows[i][5] : '';
  const row = [date, cls, present, absent.length, absent.join(', '), late, nowTime_()];
  if (i >= 0) log.sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  else log.sh.appendRow(row);
  return { ok: true, date: date };
}
