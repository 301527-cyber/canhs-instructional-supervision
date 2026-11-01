/**
 * CANHS Instructional Supervision — shared record (Google Apps Script)
 *
 * Access is controlled by the "Access" tab of this Google Sheet. Each person signs in with their
 * email address and a 4-digit passcode:
 *   email | passcode | role | department | name
 *   role = admin  → School Head: sees and edits everything, sets departments and access codes
 *          edit   → can see and edit every department (optional; leave unused to keep departments apart)
 *          dept   → department head or Master Teacher who supervises: sees and edits ONLY the records of teachers
 *                   in the department(s) named. One person may handle 1–2 departments: write them as "English, Filipino".
 *          view   → can read everything except Master Teachers' records, cannot save
 * Master Teachers' IS plans and supervisory records (observations, on-the-spot tests, lesson plan
 * monitoring, PMCFs) are sent only to the School Head (admin), never to department heads.
 * Records of a department head's teachers are filtered HERE, on the server, so other departments'
 * records are never sent to that department head's browser.
 *
 * First time: Extensions ▸ Apps Script ▸ paste this file ▸ run setup() once ▸ Deploy ▸ New deployment ▸
 * Web app (Execute as: Me, Who has access: Anyone) ▸ copy the /exec address into the app's Connection page.
 * setup() writes a School Head code in the Access tab; change it to your own.
 */
var VERSION = 'email-passcode-4';
var SHEET_REC = 'Records', SHEET_ACC = 'Access';
var COLLS = ['obs', 'coach', 'plans', 'config'];
var CELL = 45000;            // characters per cell; a longer record continues in the next cells of its row (columns G, H, …)
var MAX_PARTS = 40;

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.appendRow(header); s.setFrozenRows(1); }
  return s;
}
function recSheet_() { return sheet_(SHEET_REC, ['coll', 'id', 'department', 'json', 'updated', 'by']); }
var ACC_HEAD = ['email', 'passcode', 'role', 'department', 'name'];
function accSheet_() {
  var s = sheet_(SHEET_ACC, ACC_HEAD);
  if (String(s.getRange(1, 1).getValue()).trim().toLowerCase() !== 'email') {   // an Access tab from the code-only version
    var old = s.getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]).trim(); });
    s.clear(); s.appendRow(ACC_HEAD);
    old.forEach(function (r) { s.appendRow(['', String(r[0]).trim(), r[1], r[2], r[3]]); });   // old codes kept in the passcode column; add the emails
  }
  s.getRange('B:B').setNumberFormat('@');            // keep passcodes like 0123 as text
  return s;
}
function pin_(v) { v = String(v == null ? '' : v).trim(); return /^\d{1,4}$/.test(v) ? ('0000' + v).slice(-4) : v; }
function email_(v) { return String(v || '').trim().toLowerCase(); }

function setup() {
  recSheet_();
  var a = accSheet_();
  var hasAdmin = listCodes_().some(function (r) { return r.role === 'admin' && r.email; });
  if (!hasAdmin) { var me = ''; try { me = Session.getEffectiveUser().getEmail(); } catch (e) {} a.appendRow([me, randomPin_(), 'admin', '', 'School Head']); }
  showCodes();
}
/* Type the email and passcode you use in the app, then Run testLogin to see what the script finds */
function testLogin() {
  var EMAIL = '301527@deped.gov.ph', PASSCODE = '0000';
  CacheService.getScriptCache().remove('fail:' + email_(EMAIL));
  var a = access_(EMAIL + '|' + PASSCODE);
  Logger.log(a && a.role ? 'ACCEPTED as ' + a.role + (a.dept ? ' (' + a.dept + ')' : '') : 'NOT ACCEPTED. Check the Access tab: email in column A, passcode in column B, admin in column C.');
  showCodes();
}
/* Run this to see who can sign in, in the Execution log */
function showCodes() {
  listCodes_().forEach(function (r) {
    Logger.log((r.role === 'admin' ? 'SCHOOL HEAD' : r.role.toUpperCase() + (r.department ? ' (' + r.department + ')' : '')) + ': ' + (r.email || '(no email yet)') + '  passcode ' + r.pin);
  });
  Logger.log('Sign-ins are kept in the "Access" tab of the Google Sheet: ' + ss_().getUrl());
}
/* To set your own School Head sign-in: type your email and a 4-digit passcode between the quotes, then Run setMyLogin */
function setMyLogin() {
  var MY_EMAIL = 'your.name@deped.gov.ph', MY_PASSCODE = '0000';
  if (MY_EMAIL === 'your.name@deped.gov.ph' || !/^\d{4}$/.test(MY_PASSCODE)) { Logger.log('Type your email and a 4-digit passcode first.'); return; }
  var sh = accSheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][2]).trim().toLowerCase() === 'admin') { sh.getRange(i + 1, 1, 1, 2).setValues([[MY_EMAIL.toLowerCase(), MY_PASSCODE]]); Logger.log('SCHOOL HEAD sign-in: ' + MY_EMAIL + '  passcode ' + MY_PASSCODE); return; }
  sh.appendRow([MY_EMAIL.toLowerCase(), MY_PASSCODE, 'admin', '', 'School Head']); Logger.log('SCHOOL HEAD sign-in: ' + MY_EMAIL + '  passcode ' + MY_PASSCODE);
}
function randomPin_() { return ('000' + Math.floor(Math.random() * 10000)).slice(-4); }
function randomCode_() {
  var c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 8; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}

/* code = "email|passcode". Five wrong passcodes lock that email for 15 minutes. */
function access_(code) {
  code = String(code || '').trim();
  var k = code.lastIndexOf('|'); if (k < 1) return null;
  var email = email_(code.slice(0, k)), pin = pin_(code.slice(k + 1));
  if (!email || !pin) return null;
  var cache = CacheService.getScriptCache(), key = 'fail:' + email, fails = +(cache.get(key) || 0);
  if (fails >= 5) return { locked: true };
  var v = accSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (email_(v[i][0]) === email && pin_(v[i][1]) === pin) {
      var role = String(v[i][2]).trim().toLowerCase();
      if (['admin', 'edit', 'dept', 'view'].indexOf(role) < 0) return null;
      var dept = String(v[i][3] || '').trim();
      if (role === 'dept' && !dept) return null;          // a department head's row must name the department
      cache.remove(key);
      var depts = dept.split(/\s*[,;&\/]\s*/).filter(String);
      return { role: role, dept: depts.join(', '), depts: depts, name: String(v[i][4] || ''), email: email };
    }
  }
  cache.put(key, String(fails + 1), 900);
  return null;
}
function listCodes_() {
  var v = accSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() || String(v[i][1]).trim()) out.push({ email: email_(v[i][0]), pin: pin_(v[i][1]), role: String(v[i][2]).trim().toLowerCase(), department: String(v[i][3] || '').trim(), name: String(v[i][4] || '') });
  return out;
}

function readAll_() {
  var v = recSheet_().getDataRange().getValues(), out = {};
  COLLS.forEach(function (c) { out[c] = []; });
  for (var i = 1; i < v.length; i++) {
    var coll = String(v[i][0]), id = String(v[i][1]);
    if (!out[coll] || !id) continue;
    try { var o = JSON.parse(String(v[i][3]) + v[i].slice(6).join('')); o.id = id; out[coll].push({ o: o, row: i + 1 }); } catch (e) {}
  }
  return out;
}
/* the row of one record, found from the first two columns only (fast even when the sheet is large) */
function findRow_(sh, coll, id) {
  var n = sh.getLastRow(); if (n < 2) return 0;
  var v = sh.getRange(2, 1, n - 1, 2).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === coll && String(v[i][1]) === id) return i + 2;
  return 0;
}
function rowObj_(sh, row) {
  var r = sh.getRange(row, 1, 1, Math.max(6, sh.getLastColumn())).getValues()[0];
  try { return JSON.parse(String(r[3]) + r.slice(6).join('')); } catch (e) { return null; }
}
function assignOf_(all) {
  var s = (all.config || []).filter(function (r) { return r.o.id === 'school'; })[0];
  return (s && s.o.assign) || {};
}
/* the department a record belongs to: its own department, else the teacher's assigned department */
function deptOf_(coll, o, assign) {
  if (coll === 'config') return '';
  if (o && o.dept) return String(o.dept);
  var n = coll === 'coach' ? o.ratee : o.teacher;
  return (n && assign[n]) || '';
}

/* Master Teachers' plans and supervisory records are for the School Head only */
function mtNames_(all) {
  var s = (all.config || []).filter(function (r) { return r.o.id === 'school'; })[0];
  var m = {}; ((s && s.o.mts) || []).forEach(function (n) { m[n] = 1; }); return m;
}
function isMT_(coll, o, mts) {
  if (coll === 'config' || !o) return false;
  var tp = coll === 'plans' && (o.kind === 'tplan' || o.kind === 'srep');   // a teacher's own plan or Supervisory Report entry
  var pos = coll === 'plans' ? (tp ? o.teacherPos : o.prepPos) : coll === 'coach' ? o.rateePos : o.teacherPos;
  var who = coll === 'plans' ? (tp ? o.teacher : o.prep) : coll === 'coach' ? o.ratee : o.teacher;
  return /^\s*master\s+teacher/i.test(String(pos || '')) || !!(who && mts[who]);
}
/* is department d one of the departments this sign-in handles? */
function inDept_(a, d) { return (a.depts || [a.dept]).indexOf(String(d || '')) >= 0; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') {   // lets the app check which script is running; shows no records or passcodes
    var n = 0; try { n = listCodes_().filter(function (r) { return r.email && r.pin; }).length; } catch (err) {}
    var adm = listCodes_().some(function (r) { return r.role === 'admin' && r.email && r.pin; });
    return json_({ ok: true, version: VERSION, sheet: ss_().getName(), signins: n, hasHead: adm });
  }
  var a = access_(p.code);
  if (a && a.locked) return json_({ ok: false, error: 'locked_out' });
  if (!a) return json_({ ok: false, error: 'code' });
  if (p.action === 'codes') {
    if (a.role !== 'admin') return json_({ ok: false, error: 'admin' });
    return json_({ ok: true, codes: listCodes_() });
  }
  var all = readAll_(), assign = assignOf_(all), mts = mtNames_(all), data = {};
  COLLS.forEach(function (c) {
    data[c] = all[c].filter(function (r) {
      if (a.role === 'admin' || c === 'config') return true;
      if (isMT_(c, r.o, mts)) return false;
      if (a.role !== 'dept') return true;
      return inDept_(a, deptOf_(c, r.o, assign));
    }).map(function (r) { return r.o; });
  });
  return json_({ ok: true, role: a.role, dept: a.role === 'dept' ? a.dept : '', depts: a.role === 'dept' ? a.depts : [], name: a.name, data: data });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  var lock = null;
  try {
    var a = access_(body.code);
    if (a && a.locked) return json_({ ok: false, error: 'locked_out' });
    if (!a) return json_({ ok: false, error: 'code' });
    if (a.role === 'view') return json_({ ok: false, error: 'readonly' });
    lock = LockService.getScriptLock();
    if (!lock.tryLock(28000)) { lock = null; return json_({ ok: false, error: 'busy' }); }   // someone else is saving: the app tries again
    if (body.action === 'codes') return json_(saveCodes_(a, body.codes));
    var coll = String(body.coll || ''), id = String(body.id || '');
    if (COLLS.indexOf(coll) < 0 || !id) return json_({ ok: false, error: 'bad_request' });
    if (coll === 'config' && a.role !== 'admin') return json_({ ok: false, error: 'admin' });
    var sh = recSheet_(), srow = findRow_(sh, 'config', 'school'), school = (srow && rowObj_(sh, srow)) || {};
    var assign = school.assign || {}, mts = {}; (school.mts || []).forEach(function (n) { mts[n] = 1; });
    var row0 = findRow_(sh, coll, id), curO = row0 ? rowObj_(sh, row0) : null;
    if (a.role !== 'admin' && ((curO && isMT_(coll, curO, mts)) || (body.action === 'put' && isMT_(coll, body.data || {}, mts)))) return json_({ ok: false, error: 'mt' });
    if (a.role === 'dept' && curO && !inDept_(a, deptOf_(coll, curO, assign))) return json_({ ok: false, error: 'dept' });
    if (body.action === 'del') {
      if (row0) sh.deleteRow(row0);
      return json_({ ok: true });
    }
    if (body.action !== 'put') return json_({ ok: false, error: 'bad_request' });
    var data = body.data || {};
    var dept = deptOf_(coll, data, assign);
    if (a.role === 'dept') {
      // a record with no department yet belongs to the department of the one who saves it
      if (!dept && coll !== 'config' && a.depts.length) { dept = a.depts[0]; data.dept = dept; }
      var who = coll === 'coach' ? data.ratee : data.teacher;
      if (!inDept_(a, dept) || (who && assign[who] && !inDept_(a, assign[who]))) return json_({ ok: false, error: 'dept' });
    }
    var copy = {}; Object.keys(data).forEach(function (k) { if (k !== 'id') copy[k] = data[k]; });
    var txt = JSON.stringify(copy), parts = [];
    for (var p = 0; p < txt.length; ) { var q = Math.min(txt.length, p + CELL);
      while (q < txt.length && q > p + 1 && /[=+\-@\s]/.test(txt.charAt(q))) q--;   // a cell never starts with = + - @ (Sheets would read a formula)
      parts.push(txt.slice(p, q)); p = q; }
    if (!parts.length) parts = ['{}'];
    if (parts.length > MAX_PARTS) return json_({ ok: false, error: 'too_large' });
    var row = [coll, id, dept, parts[0], new Date(), a.email || a.name || a.role].concat(parts.slice(1));
    var width = Math.max(row.length, sh.getLastColumn());
    while (row.length < width) row.push('');      // clears the extra cells of a record that became shorter
    var at = row0 || sh.getLastRow() + 1;
    sh.getRange(at, 4).setNumberFormat('@'); if (width > 6) sh.getRange(at, 7, 1, width - 6).setNumberFormat('@');
    sh.getRange(at, 1, 1, width).setValues([row]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server', detail: String(err && err.message || err) });
  } finally { if (lock) lock.releaseLock(); }
}

/* School Head only: replace the department heads' sign-ins (rows with role dept/edit/view); School Head rows are kept */
function saveCodes_(a, codes) {
  if (a.role !== 'admin') return { ok: false, error: 'admin' };
  if (!Array.isArray(codes)) return { ok: false, error: 'bad_request' };
  var sh = accSheet_(), keep = listCodes_().filter(function (r) { return r.role === 'admin'; });
  var seen = {}; keep.forEach(function (r) { seen[r.email] = 1; });
  var rows = keep.map(function (r) { return [r.email, r.pin, r.role, r.department, r.name]; });
  for (var i = 0; i < codes.length; i++) {
    var c = codes[i] || {}, email = email_(c.email), pin = pin_(c.pin), role = String(c.role || 'dept').toLowerCase();
    if (!email && !pin) continue;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'email' };
    if (!/^\d{4}$/.test(pin)) return { ok: false, error: 'pin' };
    if (seen[email]) return { ok: false, error: 'duplicate' };
    if (['dept', 'edit', 'view'].indexOf(role) < 0) continue;
    seen[email] = 1; rows.push([email, pin, role, String(c.department || ''), String(c.name || '')]);
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
  return { ok: true };
}
