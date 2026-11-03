/**
 * Learnovate Qatar · CELPIP Prep Course — backend
 * Google Apps Script bound to a Google Sheet. The Sheet is the database:
 *   Batches  · one row per class batch
 *   Students · one row per student, with the access code you assign
 *   Progress · each student's saved practice (JSON, split across p1–p8)
 *   Feedback · your feedback to each student
 *
 *   Messages · messages and email-writing tasks between you and students
 *
 * Setup: see README.md. Set the script property TEACHER_PASSWORD before use.
 * Run testEmail once from the editor so Google asks for permission to send email.
 */
const SHEET_HEADERS = {
  Batches:  ['id', 'name', 'start', 'schedule', 'createdAt'],
  Students: ['code', 'name', 'email', 'phone', 'batch', 'createdAt', 'updatedAt'],
  Progress: ['code', 'updatedAt', 'p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8'],
  Feedback: ['code', 'text', 'updatedAt'],
  Messages: ['id', 'code', 'batch', 'from', 'kind', 'subject', 'body', 'createdAt', 'readS', 'readT', 'emailed', 'replyTo']
};
const CHUNK = 45000;          // a Sheets cell holds at most 50,000 characters
const CODE_RE = /^[A-Z0-9-]{4,20}$/;

function doGet() {
  return out_({ ok: true, service: 'Learnovate Qatar CELPIP Prep Course API' });
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

/** Run once from the editor to create the sheets. */
function setup() {
  Object.keys(SHEET_HEADERS).forEach(sheet_);
}

/** Run once from the editor: sends you a test email and grants email permission. */
function testEmail() {
  const me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(me, 'Learnovate Qatar CELPIP: email is working', 'Your course can now email students. Daily quota left: ' + MailApp.getRemainingDailyQuota());
}

function route_(q) {
  const a = String(q.action || '');
  if (a === 'ping') return {};

  if (a.indexOf('student.') === 0) {
    const st = find_('Students', 'code', normCode_(q.code));
    if (!st) fail_('bad_code', 'Access code not found');
    if (a === 'student.save') { saveProgress_(st.code, q.progress); return { savedAt: Date.now() }; }
    if (a === 'student.messages') {
      const fb0 = find_('Feedback', 'code', st.code);
      return { messages: studentMsgs_(st), feedback: fb0 ? { text: String(fb0.text), updatedAt: Number(fb0.updatedAt) || 0 } : null };
    }
    if (a === 'student.send') {
      const body = clean_(q.body, 6000);
      if (!body) fail_('bad_request', 'Message is empty');
      const row = { id: newId_('M'), code: st.code, batch: String(st.batch || ''), from: 'student',
        kind: q.kind === 'task-reply' ? 'task-reply' : 'message', subject: clean_(q.subject, 150), body: body,
        createdAt: Date.now(), readS: Date.now(), readT: '', emailed: '', replyTo: clean_(q.replyTo, 20) };
      write_('Messages', row);
      notifyTeacher_(st, row);
      return { message: pubMsg_(row) };
    }
    if (a === 'student.read') { markRead_('readS', q.ids, function (r) { return String(r.code) === st.code; }); return {}; }
    if (a !== 'student.login') fail_('bad_request', 'Unknown action');
    const b = st.batch ? find_('Batches', 'id', st.batch) : null;
    const fb = find_('Feedback', 'code', st.code);
    return {
      student: pubStudent_(st),
      batch: b ? pubBatch_(b) : null,
      progress: readProgress_(find_('Progress', 'code', st.code)),
      feedback: fb ? { text: String(fb.text), updatedAt: Number(fb.updatedAt) || 0 } : null,
      messages: studentMsgs_(st)
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
          const col = SHEET_HEADERS.Messages.indexOf('code') + 1;
          rows_('Messages').forEach(function (r) { if (String(r.code) === orig) sheet_('Messages').getRange(r._row, col).setValue(code); });
        }
        return { student: pubStudent_(row) };
      }

      case 'teacher.student.delete': {
        const code = normCode_(q.code);
        ['Students', 'Progress', 'Feedback'].forEach(function (n) { del_(n, 'code', code); });
        rows_('Messages').filter(function (r) { return String(r.code) === code; }).reverse()
          .forEach(function (r) { sheet_('Messages').deleteRow(r._row); });
        return {};
      }

      case 'teacher.send': {
        const to = q.to || {};
        const body = clean_(q.body, 6000);
        if (!body) fail_('bad_request', 'Message is empty');
        const kind = ['message', 'task', 'invite'].indexOf(q.kind) >= 0 ? q.kind : 'message';
        let recipients = [], row;
        if (to.code) {
          const st = find_('Students', 'code', normCode_(to.code));
          if (!st) fail_('bad_code', 'Student not found');
          recipients = [st];
          row = { code: st.code, batch: String(st.batch || '') };
        } else if (to.batch) {
          if (!find_('Batches', 'id', to.batch)) fail_('bad_request', 'Batch not found');
          recipients = rows_('Students').filter(function (r) { return String(r.batch) === String(to.batch); });
          row = { code: '', batch: String(to.batch) };
        } else fail_('bad_request', 'No recipient');
        Object.assign(row, { id: newId_('M'), from: 'teacher', kind: kind, subject: clean_(q.subject, 150), body: body,
          createdAt: Date.now(), readS: '', readT: Date.now(), emailed: '', replyTo: clean_(q.replyTo, 20) });
        const result = { sent: 0, skipped: 0, failed: 0 };
        if (q.email) {
          recipients.forEach(function (st) {
            if (!st.email) { result.skipped++; return; }
            try {
              MailApp.sendEmail({ to: String(st.email), name: 'Learnovate Qatar · CELPIP Prep Course',
                subject: row.subject || (kind === 'task' ? 'New CELPIP email-writing task' : 'Message from your CELPIP teacher'),
                body: emailBody_(row, st, clean_(q.siteUrl, 300)) });
              result.sent++;
            } catch (err) { result.failed++; }
          });
          row.emailed = result.sent ? new Date().toISOString() : '';
        }
        write_('Messages', row);
        return { message: pubMsg_(row), email: result };
      }

      case 'teacher.read':
        markRead_('readT', q.ids, function () { return true; });
        return {};

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
    progress: progress, feedback: feedback,
    messages: rows_('Messages').map(pubMsg_)
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

/* ---------- messages ---------- */
function pubMsg_(r) {
  return { id: String(r.id), code: String(r.code || ''), batch: String(r.batch || ''), from: String(r.from || ''),
    kind: String(r.kind || 'message'), subject: String(r.subject || ''), body: String(r.body || ''),
    createdAt: Number(r.createdAt) || 0, readS: Number(r.readS) || 0, readT: Number(r.readT) || 0,
    emailed: String(r.emailed || ''), replyTo: String(r.replyTo || '') };
}
function studentMsgs_(st) {
  return rows_('Messages').filter(function (r) {
    return String(r.code) === st.code || (String(r.code) === '' && st.batch && String(r.batch) === String(st.batch));
  }).map(pubMsg_);
}
function markRead_(col, ids, allowed) {
  if (!Array.isArray(ids) || !ids.length) return;
  const want = {}; ids.forEach(function (i) { want[String(i)] = true; });
  const c = SHEET_HEADERS.Messages.indexOf(col) + 1, now = Date.now(), sh = sheet_('Messages');
  rows_('Messages').forEach(function (r) { if (want[String(r.id)] && !r[col] && allowed(r)) sh.getRange(r._row, c).setValue(now); });
}
function emailBody_(row, st, siteUrl) {
  let t = 'Hi ' + (st.name || 'there') + ',\n\n' + row.body + '\n\n';
  if (row.kind === 'task') t += 'This is a CELPIP email-writing task. Write your reply (150–200 words) on your course dashboard, under Messages.\n';
  else t += 'You can reply on your course dashboard, under Messages.\n';
  if (siteUrl) t += siteUrl + '\n';
  if (row.kind === 'invite') t += 'Your access code: ' + st.code + '\n';
  return t + '\nLearnovate Qatar · CELPIP Prep Course';
}
function notifyTeacher_(st, row) {
  if (PropertiesService.getScriptProperties().getProperty('NOTIFY_ME') !== 'yes') return;
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'New message from ' + st.name + (row.subject ? ': ' + row.subject : ''),
      row.body + '\n\nOpen the Teacher Portal to reply.');
  } catch (e) {}
}
function newId_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/* ---------- sheet helpers ---------- */
function fail_(code, msg) { const e = new Error(msg || code); e.code = code; throw e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@');   // keep codes, phones and dates as plain text
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
