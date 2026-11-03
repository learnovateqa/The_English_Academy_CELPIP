/**
 * CELPIP Prep Course — backend
 * Google Apps Script bound to a Google Sheet. The Sheet is the database:
 *   Batches  · one row per class batch
 *   Students · one row per student, with the access code you assign
 *   Progress · each student's saved practice (JSON, split across p1–p8)
 *   Feedback · your feedback to each student
 *
 * Setup: see README.md. Set the script property TEACHER_PASSWORD before use.
 */
const SHEET_HEADERS = {
  Batches:  ['id', 'name', 'start', 'schedule', 'createdAt'],
  Students: ['code', 'name', 'email', 'phone', 'batch', 'createdAt', 'updatedAt'],
  Progress: ['code', 'updatedAt', 'p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8'],
  Feedback: ['code', 'text', 'updatedAt']
};
const CHUNK = 45000;          // a Sheets cell holds at most 50,000 characters
const CODE_RE = /^[A-Z0-9-]{4,20}$/;

function doGet() {
  return out_({ ok: true, service: 'CELPIP Prep Course API' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return out_({ ok: false, error: 'busy' }); }
  try { return out_(Object.assign({ ok: true }, route_(req))); }
  catch (err) { return out_({ ok: false, error: err.code || 'server_error', message: String((err && err.message) || err) }); }
  finally { lock.releaseLock(); }
}

/** Run once from the editor to create the four sheets. */
function setup() {
  Object.keys(SHEET_HEADERS).forEach(sheet_);
}

function route_(q) {
  const a = String(q.action || '');
  if (a === 'ping') return {};

  if (a === 'student.login' || a === 'student.save') {
    const st = find_('Students', 'code', normCode_(q.code));
    if (!st) fail_('bad_code', 'Access code not found');
    if (a === 'student.save') { saveProgress_(st.code, q.progress); return { savedAt: Date.now() }; }
    const b = st.batch ? find_('Batches', 'id', st.batch) : null;
    const fb = find_('Feedback', 'code', st.code);
    return {
      student: pubStudent_(st),
      batch: b ? pubBatch_(b) : null,
      progress: readProgress_(find_('Progress', 'code', st.code)),
      feedback: fb ? { text: String(fb.text), updatedAt: Number(fb.updatedAt) || 0 } : null
    };
  }

  if (a.indexOf('teacher.') === 0) {
    const pass = PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD');
    if (!pass) fail_('no_password', 'Set TEACHER_PASSWORD in Project Settings > Script properties');
    if (String(q.pass || '') !== pass) fail_('bad_password', 'Wrong teacher password');

    switch (a) {
      case 'teacher.login': return {};
      case 'teacher.data': return teacherData_();

      case 'teacher.batch.save': {
        const b = q.batch || {};
        const name = clean_(b.name, 60);
        if (!name) fail_('bad_request', 'Batch name is required');
        const existing = b.id ? find_('Batches', 'id', b.id) : null;
        const row = {
          id: existing ? existing.id : 'B' + Utilities.getUuid().replace(/-/g, '').slice(0, 8),
          name: name, start: clean_(b.start, 20), schedule: clean_(b.schedule, 60),
          createdAt: existing ? existing.createdAt : Date.now()
        };
        write_('Batches', row, existing && existing._row);
        return { batch: pubBatch_(row) };
      }

      case 'teacher.batch.delete':
        del_('Batches', 'id', q.id);
        return {};

      case 'teacher.student.save': {
        const s = q.student || {};
        const code = normCode_(s.code);
        const orig = normCode_(q.originalCode);
        const name = clean_(s.name, 80);
        if (!name) fail_('bad_request', 'Student name is required');
        if (!CODE_RE.test(code)) fail_('bad_code_format', 'Codes use 4–20 letters, numbers or dashes');
        if (code !== orig && find_('Students', 'code', code)) fail_('code_taken', 'Code already in use');
        const existing = orig ? find_('Students', 'code', orig) : null;
        const row = {
          code: code, name: name, email: clean_(s.email, 120), phone: clean_(s.phone, 40), batch: clean_(s.batch, 20),
          createdAt: existing ? existing.createdAt : Date.now(), updatedAt: Date.now()
        };
        write_('Students', row, existing && existing._row);
        if (existing && orig !== code) {
          ['Progress', 'Feedback'].forEach(function (n) {
            const r = find_(n, 'code', orig);
            if (r) sheet_(n).getRange(r._row, 1).setValue(code);
          });
        }
        return { student: pubStudent_(row) };
      }

      case 'teacher.student.delete': {
        const code = normCode_(q.code);
        ['Students', 'Progress', 'Feedback'].forEach(function (n) { del_(n, 'code', code); });
        return {};
      }

      case 'teacher.feedback': {
        const code = normCode_(q.code);
        if (!find_('Students', 'code', code)) fail_('bad_code', 'Student not found');
        const existing = find_('Feedback', 'code', code);
        write_('Feedback', { code: code, text: clean_(q.text, 4000), updatedAt: Date.now() }, existing && existing._row);
        return {};
      }
    }
  }
  fail_('bad_request', 'Unknown action');
}

function teacherData_() {
  const progress = {};
  rows_('Progress').forEach(function (r) { const p = readProgress_(r); if (p) progress[String(r.code)] = p; });
  const feedback = {};
  rows_('Feedback').forEach(function (r) { feedback[String(r.code)] = { text: String(r.text), updatedAt: Number(r.updatedAt) || 0 }; });
  return {
    batches: rows_('Batches').map(pubBatch_),
    students: rows_('Students').map(pubStudent_),
    progress: progress, feedback: feedback
  };
}

function saveProgress_(code, p) {
  if (!p || typeof p !== 'object') fail_('bad_request', 'Missing progress');
  const json = JSON.stringify({
    tasks: p.tasks || {}, res: p.res || {}, targets: p.targets || {},
    log: Array.isArray(p.log) ? p.log.slice(0, 200) : [],
    updatedAt: Number(p.updatedAt) || Date.now()
  });
  if (json.length > CHUNK * 8) fail_('too_large', 'Progress is too large');
  const row = { code: code, updatedAt: Date.now() };
  for (let i = 0; i < 8; i++) row['p' + (i + 1)] = json.slice(i * CHUNK, (i + 1) * CHUNK);
  const existing = find_('Progress', 'code', code);
  write_('Progress', row, existing && existing._row);
}

function readProgress_(r) {
  if (!r) return null;
  let s = '';
  for (let i = 1; i <= 8; i++) s += String(r['p' + i] || '');
  try { return s ? JSON.parse(s) : null; } catch (e) { return null; }
}

function pubStudent_(r) {
  return { code: String(r.code), name: String(r.name || ''), email: String(r.email || ''), phone: String(r.phone || ''),
           batch: String(r.batch || ''), createdAt: Number(r.createdAt) || 0 };
}
function pubBatch_(r) {
  return { id: String(r.id), name: String(r.name || ''), start: String(r.start || ''), schedule: String(r.schedule || ''),
           createdAt: Number(r.createdAt) || 0 };
}

/* ---------- sheet helpers ---------- */
function fail_(code, msg) { const e = new Error(msg || code); e.code = code; throw e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:J').setNumberFormat('@');   // keep codes, phones and dates as plain text
    sh.appendRow(SHEET_HEADERS[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const h = v[0] || [];
  return v.slice(1).map(function (r, i) { const o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
function find_(name, key, val) {
  if (val === undefined || val === null || val === '') return null;
  return rows_(name).find(function (r) { return String(r[key]) === String(val); }) || null;
}
function write_(name, obj, rowIdx) {
  const h = SHEET_HEADERS[name];
  const vals = h.map(function (k) { return obj[k] === undefined ? '' : obj[k]; });
  const sh = sheet_(name);
  if (rowIdx) sh.getRange(rowIdx, 1, 1, h.length).setValues([vals]); else sh.appendRow(vals);
}
function del_(name, key, val) { const r = find_(name, key, val); if (r) sheet_(name).deleteRow(r._row); }
function normCode_(c) { return String(c || '').trim().toUpperCase(); }
function clean_(s, max) { return String(s === null || s === undefined ? '' : s).trim().slice(0, max || 200); }
