/**
 * 3F3 жеке кездесу — Google Sheets сервері.
 * Google Sheets → Кеңейтімдер (Extensions) → Apps Script ішіне қойыңыз.
 * Төмендегі құпиясөзді міндетті түрде өзгертіңіз.
 */
const ADMIN_PASSWORD = 'ОСЫНЫ-ӨЗГЕРТІҢІЗ';

const SH = {
  students: ['id', 'token', 'name', 'contact', 'createdAt', 'updatedAt', 'answers', 'ikigai'],
  slots: ['slot', 'studentId', 'at'],
  notes: ['studentId', 'updatedAt', 'data'],
  config: ['key', 'value']
};
const READABLE = 'Жауаптар';
const MAX = 45000;

function doGet() { return out({ ok: true, app: '3F3 жеке кездесу API' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'bad_json' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return out(Object.assign({ ok: true }, route(req)));
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function route(r) {
  switch (r.action) {
    case 'register': return register(r);
    case 'me': return me(student(r.token));
    case 'save': return save(student(r.token), r);
    case 'book': return book(student(r.token), r.slot);
    case 'unbook': { const s = student(r.token); clearSlots(s.id); return { slots: slotsFor(s.id) }; }
    case 'admin': admin(r.pw); return everything();
    case 'saveNote': admin(r.pw); return saveNote(r.id, r.note);
    case 'free': admin(r.pw); clearSlot(r.slot); return {};
    case 'config': admin(r.pw); return { config: setConfig(r.config || {}) };
    default: throw new Error('unknown_action');
  }
}

/* ---------- sheets ---------- */
function sheet(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(SH[name]); s.setFrozenRows(1); }
  return s;
}
function rows(name) {
  const v = sheet(name).getDataRange().getValues();
  const h = v.shift();
  return v.map((row, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = row[j]); return o; });
}
function str(v, n) { return String(v == null ? '' : v).slice(0, n || 2000); }
function json(v) { try { return JSON.parse(v || '{}'); } catch (e) { return {}; } }
function setup() { Object.keys(SH).forEach(sheet); }

/* ---------- auth ---------- */
function admin(pw) {
  if (ADMIN_PASSWORD === 'ОСЫНЫ-ӨЗГЕРТІҢІЗ') throw new Error('change_password');
  if (pw !== ADMIN_PASSWORD) { Utilities.sleep(800); throw new Error('bad_password'); }
}
function student(token) {
  const t = str(token, 64);
  const s = t && rows('students').find(r => r.token === t);
  if (!s) throw new Error('not_found');
  return s;
}

/* ---------- config & slots ---------- */
function getConfig() {
  const c = { slotMin: 20, lunch: true, closed: false };
  rows('config').forEach(r => { try { c[r.key] = JSON.parse(r.value); } catch (e) {} });
  return c;
}
function setConfig(c) {
  const s = sheet('config'), have = rows('config');
  ['slotMin', 'lunch', 'closed'].forEach(k => {
    if (!(k in c)) return;
    const v = k === 'slotMin' ? Math.min(60, Math.max(10, Number(c[k]) || 20)) : !!c[k];
    const r = have.find(x => x.key === k);
    if (r) s.getRange(r._row, 2).setValue(JSON.stringify(v)); else s.appendRow([k, JSON.stringify(v)]);
  });
  return getConfig();
}
function slotsFor(myId) {
  const o = {};
  rows('slots').forEach(r => { if (r.slot) o[r.slot] = { uid: r.studentId === myId ? myId : 'x' }; });
  return o;
}
function clearSlots(studentId) {
  const s = sheet('slots');
  rows('slots').filter(r => r.studentId === studentId).reverse().forEach(r => s.deleteRow(r._row));
}
function clearSlot(slot) {
  const s = sheet('slots');
  rows('slots').filter(r => r.slot === slot).reverse().forEach(r => s.deleteRow(r._row));
}
function book(st, slot) {
  if (!/^s\d{4}$/.test(String(slot))) throw new Error('bad_slot');
  if (getConfig().closed) throw new Error('closed');
  const taken = rows('slots').find(r => r.slot === slot);
  if (taken && taken.studentId !== st.id) throw new Error('taken');
  if (!taken) { clearSlots(st.id); sheet('slots').appendRow([slot, st.id, new Date().toISOString()]); }
  return { slots: slotsFor(st.id) };
}

/* ---------- students ---------- */
function register(r) {
  const name = str(r.name, 120).trim();
  if (!name) throw new Error('name_required');
  const id = Utilities.getUuid().slice(0, 8);
  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  const now = new Date().toISOString();
  sheet('students').appendRow([id, token, name, str(r.contact, 120), now, now, '{}', '{}']);
  return { id: id, token: token };
}
function pub(s) {
  return { id: s.id, name: s.name, contact: s.contact, createdAt: s.createdAt, updatedAt: s.updatedAt, answers: json(s.answers), ikigai: json(s.ikigai) };
}
function me(s) {
  const n = rows('notes').find(r => r.studentId === s.id);
  return { me: pub(s), note: n ? json(n.data) : null, slots: slotsFor(s.id), config: getConfig() };
}
function save(s, r) {
  const a = JSON.stringify(r.answers || {}), k = JSON.stringify(r.ikigai || {});
  if (a.length > MAX || k.length > MAX) throw new Error('too_long');
  const now = new Date().toISOString();
  sheet('students').getRange(s._row, 6, 1, 3).setValues([[now, a, k]]);
  if (Array.isArray(r.flat)) readable(s, r.flat);
  return {};
}
// Мұғалім Sheets ішінде де оқи алуы үшін: бір студент — бір жол, бір сұрақ — бір баған.
function readable(s, flat) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(READABLE);
  const header = ['Аты-жөні', 'Байланыс', 'Жаңартылды'].concat(flat.map(p => str(p[0], 300)));
  if (!sh) { sh = ss.insertSheet(READABLE); sh.appendRow(header); sh.setFrozenRows(1); sh.setFrozenColumns(1); }
  else sh.getRange(1, 1, 1, header.length).setValues([header]);
  const row = [s.name, s.contact, new Date()].concat(flat.map(p => str(p[1], 5000)));
  const ids = sh.getRange(1, header.length + 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(x => x[0]);
  let at = ids.indexOf(s.id) + 1;
  if (!at) { at = sh.getLastRow() + 1; }
  sh.getRange(at, 1, 1, row.length).setValues([row]);
  sh.getRange(at, header.length + 1).setValue(s.id);
}

/* ---------- teacher ---------- */
function everything() {
  const students = {}, notes = {}, slots = {};
  rows('students').forEach(s => { if (s.id) students[s.id] = pub(s); });
  rows('notes').forEach(n => { if (n.studentId) notes[n.studentId] = json(n.data); });
  rows('slots').forEach(r => { if (r.slot) slots[r.slot] = { uid: r.studentId }; });
  return { students: students, notes: notes, slots: slots, config: getConfig() };
}
function saveNote(id, note) {
  const data = JSON.stringify(note || {});
  if (data.length > MAX) throw new Error('too_long');
  const s = sheet('notes'), n = rows('notes').find(r => r.studentId === id);
  const now = new Date().toISOString();
  if (n) s.getRange(n._row, 2, 1, 2).setValues([[now, data]]); else s.appendRow([id, now, data]);
  return {};
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
