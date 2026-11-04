/**
 * INTERCENTROS 10° EDICIÓN · Servidor de datos (Google Apps Script)
 * Todos los datos (usuarios de centros, jueces, calendario, puntuaciones,
 * clasificación y registro de actividad) se guardan en una hoja de cálculo
 * de tu Google Drive. La app (index.html, alojada donde quieras) se conecta
 * a la URL /exec de esta implementación, que el administrador introduce en
 * el primer acceso.
 *
 * Implementar como Aplicación web: Ejecutar como «Yo» · Acceso «Cualquier usuario».
 */

const TITLE = 'INTERCENTROS 10° EDICIÓN';

// Secuencia oficial de disciplinas (Fútbol 7 abre; Datch-ball cierra en Mixto)
const EVENTS = [
  {id:'f7', n:'Fútbol 7',   c:['M','F']},
  {id:'vo', n:'Vóley',      c:['M','F']},
  {id:'pa', n:'Pádel',      c:['M','F']},
  {id:'ho', n:'Hockey',     c:['M','F']},
  {id:'ba', n:'Bádminton',  c:['M','F']},
  {id:'at', n:'Atletismo',  c:['M','F']},
  {id:'es', n:'Escalada',   c:['M','F']},
  {id:'da', n:'Datch-ball', c:['Mixto']}
];
const CATN = {M:'Masculino', F:'Femenino', Mixto:'Mixto'};
const COLORS = ['#dc2626','#2563eb','#16a34a','#eab308','#9333ea','#ea580c','#0891b2','#db2777','#65a30d','#0d9488','#4f46e5','#475569'];

const SH = {CONFIG:'Config', CENTROS:'Centros', JUECES:'Jueces', ORG:'Organización', ASIG:'Asignaciones', CAL:'Calendario', PUNT:'Puntuaciones', CLAS:'Clasificación', REG:'Registro'};
const HEAD = {
  'Config':       ['clave','valor'],
  'Centros':      ['id','nombre','color','usuario','contraseña'],
  'Jueces':       ['usuario','contraseña'],
  'Organización': ['usuario','contraseña'],
  'Asignaciones': ['disciplina_id','disciplina','juez'],
  'Calendario':   ['disciplina_id','actividad','categoría','fecha','inicio','fin','lugar'],
  'Puntuaciones': ['prueba','disciplina','categoría','centro_id','centro','posición','puntos','registrado_por','actualizado'],
  'Registro':     ['fecha','usuario','rol','acción','detalle']
};
const TEXT_SHEETS = ['Config','Centros','Jueces','Organización','Asignaciones','Calendario'];
const SESSION_TTL = 21600; // 6 h
const ERR_SESSION = {error:'Tu sesión ha caducado. Vuelve a entrar.', code:'session'};

// Borrador de calendario (editable desde la app o en la pestaña "Calendario")
const DRAFT_CALENDAR = [
  ['f7','Fútbol 7','M','','09:00','10:15','Campo A'],
  ['f7','Fútbol 7','F','','09:00','10:15','Campo B'],
  ['vo','Vóley','M','','10:30','11:30','Pista 1'],
  ['vo','Vóley','F','','10:30','11:30','Pista 2'],
  ['pa','Pádel','M','','11:45','12:45','Pistas de pádel 1–3'],
  ['pa','Pádel','F','','11:45','12:45','Pistas de pádel 4–6'],
  ['ho','Hockey','M','','13:00','14:00','Campo A'],
  ['ho','Hockey','F','','13:00','14:00','Campo B'],
  ['','Comida y barbacoa','','','14:00','15:15','Zona de picnic'],
  ['ba','Bádminton','M','','15:15','16:00','Pabellón, pistas 1–3'],
  ['ba','Bádminton','F','','15:15','16:00','Pabellón, pistas 4–6'],
  ['at','Atletismo','M','','16:15','17:00','Pista de atletismo'],
  ['at','Atletismo','F','','16:15','17:00','Pista de atletismo'],
  ['es','Escalada','M','','17:15','17:45','Rocódromo'],
  ['es','Escalada','F','','17:45','18:15','Rocódromo'],
  ['da','Datch-ball','Mixto','','18:30','19:15','Pabellón'],
  ['','Podio y entrega de premios','','','19:30','20:00','Pabellón']
];

/* ---------- Web app ---------- */
const URL_RE = /^https:\/\/script\.google\.com\/(a\/macros\/[^\/]+\/|macros\/)s\/[\w-]+\/exec$/;

// Funciones que la app puede llamar
const API = {getPublic: getPublic, login: login, logout: logout, getState: getState, saveScores: saveScores,
  saveCenters: saveCenters, saveJudges: saveJudges, saveOrgs: saveOrgs, saveAssignments: saveAssignments, saveCalendar: saveCalendar, saveSettings: saveSettings, resetScores: resetScores};

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Abrir la URL en el navegador sirve para comprobar que el servidor responde
function doGet() {
  ss();
  return json({app:'intercentros', title:TITLE, ok:true});
}

// La app envía {fn, args} como texto plano (evita la verificación CORS previa)
function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const f = Object.prototype.hasOwnProperty.call(API, b.fn) ? API[b.fn] : null;
    out = f ? f.apply(null, Array.isArray(b.args) ? b.args : []) : {error:'Función desconocida.'};
  } catch (err) {
    out = {error:'Error del servidor: ' + (err && err.message || err)};
  }
  return json(out);
}

function movedTo() { return PropertiesService.getScriptProperties().getProperty('MOVED_TO') || null; }

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Intercentros')
    .addItem('Preparar / reparar hojas', 'setup')
    .addToUi();
}

// Las ediciones hechas a mano en la hoja se reflejan en la app
function onEdit() {
  try { bump(); } catch (e) {}
}

/** Ejecútalo una vez desde el editor para autorizar y crear las pestañas. */
function setup() {
  const s = ss();
  rebuildRanking(s);
  Logger.log('Hoja lista: ' + s.getUrl());
  return s.getUrl();
}

/* ---------- Hoja de cálculo ---------- */
let _ss = null;
function ss() {
  if (_ss) return _ss;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  let s = null;
  if (id) { try { s = SpreadsheetApp.openById(id); } catch (e) { s = null; } }
  if (!s) {
    s = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('INTERCENTROS 10ª edición · datos');
    p.setProperty('SHEET_ID', s.getId());
  }
  ensure(s);
  _ss = s;
  return s;
}

function ensure(s) {
  const names = new Set(s.getSheets().map(x => x.getName()));
  let created = false;
  Object.keys(HEAD).forEach(n => {
    if (names.has(n)) return;
    const sh = s.insertSheet(n);
    if (TEXT_SHEETS.indexOf(n) >= 0) sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    seed(n, sh);
    created = true;
  });
  if (!names.has(SH.CLAS)) { s.insertSheet(SH.CLAS).setFrozenRows(1); created = true; }
  if (created) {
    s.getSheets().forEach(sh => {
      const n = sh.getName();
      if (!HEAD[n] && n !== SH.CLAS && sh.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(sh);
    });
  }
}

function seed(name, sh) {
  let data = [];
  if (name === SH.CONFIG) data = [['admin_usuario','admin'],['admin_contraseña','admin123'],['orden_estricto','sí'],['logo',''],['centros_listos','no']];
  if (name === SH.CENTROS) data = Array.from({length:12}, (_, k) => ['c' + k, 'Centro ' + (k + 1), COLORS[k % COLORS.length], 'centro' + (k + 1), '1234']);
  if (name === SH.JUECES) data = [['juez','juez123']];
  if (name === SH.ORG) data = [['organizacion','org123']];
  if (name === SH.CAL) data = DRAFT_CALENDAR;
  if (data.length) sh.getRange(2, 1, data.length, data[0].length).setValues(data);
}

function str(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt || 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function yes(v) { return ['sí','si','true','1','yes','x'].indexOf(String(v).trim().toLowerCase()) >= 0; }

function rows(s, name) {
  const sh = s.getSheetByName(name);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, Math.max(1, sh.getLastColumn())).getValues();
}
function writeRows(s, name, data) {
  const sh = s.getSheetByName(name);
  const lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, sh.getMaxColumns()).clearContent();
  if (data.length) sh.getRange(2, 1, data.length, data[0].length).setValues(data);
}

function cfgAll(s) {
  const o = {};
  rows(s, SH.CONFIG).forEach(r => { if (r[0] !== '') o[String(r[0])] = str(r[1]); });
  return o;
}
function cfgSet(s, patch) {
  const sh = s.getSheetByName(SH.CONFIG);
  const v = sh.getDataRange().getValues();
  Object.keys(patch).forEach(k => {
    const i = v.findIndex((r, j) => j > 0 && String(r[0]) === k);
    if (i < 0) { sh.appendRow([k, String(patch[k])]); v.push([k, patch[k]]); }
    else sh.getRange(i + 1, 2).setValue(String(patch[k]));
  });
}

function centers(s) {
  return rows(s, SH.CENTROS).filter(r => r[0] !== '').map(r => ({
    id: str(r[0]), name: str(r[1]), color: str(r[2]) || '#475569', user: str(r[3]).trim(), pass: str(r[4])
  }));
}
function judges(s) {
  return rows(s, SH.JUECES).filter(r => r[0] !== '').map(r => ({user: str(r[0]).trim(), pass: str(r[1])}));
}
function orgs(s) {
  return rows(s, SH.ORG).filter(r => r[0] !== '').map(r => ({user: str(r[0]).trim(), pass: str(r[1])}));
}
// Jueces asignados a cada disciplina: {f7:['juez1'], vo:['juez2','juez3'], ...}
function assignments(s) {
  const m = {};
  rows(s, SH.ASIG).forEach(r => {
    const id = str(r[0]), u = str(r[2]).trim();
    if (id && u) (m[id] = m[id] || []).push(u);
  });
  return m;
}
function calendar(s) {
  return rows(s, SH.CAL).filter(r => r.join('') !== '').map(r => ({
    id: str(r[0]), name: str(r[1]), cat: str(r[2]), date: str(r[3], 'yyyy-MM-dd'),
    start: str(r[4], 'HH:mm'), end: str(r[5], 'HH:mm'), place: str(r[6])
  }));
}
function scores(s) {
  const m = {};
  rows(s, SH.PUNT).forEach(r => {
    const k = str(r[0]), cid = str(r[3]), pos = Number(r[5]);
    if (!k || !cid || !pos) return;
    (m[k] = m[k] || {})[cid] = pos;
  });
  return m;
}

/* ---------- Versión (sincronización entre dispositivos) ---------- */
function ver() { return Number(PropertiesService.getScriptProperties().getProperty('VERSION') || 1); }
function bump() { const v = ver() + 1; PropertiesService.getScriptProperties().setProperty('VERSION', String(v)); return v; }

/* ---------- Sesiones ---------- */
function newSession(info) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, JSON.stringify(info), SESSION_TTL);
  return t;
}
function sess(t) {
  if (!t) return null;
  const c = CacheService.getScriptCache();
  const v = c.get('s_' + t);
  if (!v) return null;
  c.put('s_' + t, v, SESSION_TTL);
  return JSON.parse(v);
}
function locked(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(20000)) return {error:'La hoja está ocupada. Vuelve a intentarlo en unos segundos.'};
  try { return fn(); } finally { l.releaseLock(); }
}
function asAdmin(t, fn) {
  const info = sess(t);
  if (!info) return ERR_SESSION;
  if (info.role !== 'admin') return {error:'Solo el administrador puede hacer este cambio.'};
  return locked(() => fn(ss(), info));
}
// Administrador u organización (labores operativas)
function asStaff(t, fn) {
  const info = sess(t);
  if (!info) return ERR_SESSION;
  if (info.role !== 'admin' && info.role !== 'organizacion') return {error:'Tu usuario no puede hacer este cambio.'};
  return locked(() => fn(ss(), info));
}
function log(s, info, action, detail) {
  s.getSheetByName(SH.REG).appendRow([new Date(), info.user || '', info.role || '', action, detail || '']);
}

function state(s, info) {
  const c = cfgAll(s);
  const staff = info.role === 'admin' || info.role === 'organizacion';
  const st = {
    centers: centers(s).map(x => staff ? x : {id:x.id, name:x.name, color:x.color}),
    calendar: calendar(s),
    scores: scores(s),
    settings: {strictOrder: yes(c.orden_estricto), logo: c.logo || null, centersReady: yes(c.centros_listos)},
    version: ver()
  };
  if (staff) st.judges = judges(s);
  if (info.role !== 'centro') st.assignments = assignments(s);
  if (info.role === 'admin') {
    st.orgs = orgs(s);
    st.admin = {user: c.admin_usuario || '', pass: c.admin_contraseña || ''};
    st.sheetUrl = s.getUrl();
  }
  return st;
}

/* ---------- API para la app (google.script.run) ---------- */
function getPublic() {
  const c = cfgAll(ss());
  return {app:'intercentros', logo: c.logo || null, title: TITLE, moved: movedTo()};
}

function login(u, p) {
  u = String(u || '').trim(); p = String(p || '');
  const s = ss();
  const c = cfgAll(s);
  let info = null;
  if (u && u === c.admin_usuario && p === c.admin_contraseña) info = {role:'admin', name:'Administrador', user:u};
  else if (orgs(s).some(o => o.user === u && o.pass === p)) info = {role:'organizacion', name:'Organización', user:u};
  else if (judges(s).some(j => j.user === u && j.pass === p)) info = {role:'juez', name:u, user:u};
  else {
    const ce = centers(s).find(x => x.user === u && x.pass === p);
    if (ce) info = {role:'centro', name:ce.name, id:ce.id, user:u};
  }
  if (!info) { Utilities.sleep(700); return {error:'Usuario o contraseña incorrectos.'}; }
  log(s, info, 'Inicio de sesión', '');
  return {token: newSession(info), session: info, state: state(s, info)};
}

function logout(t) { if (t) CacheService.getScriptCache().remove('s_' + t); return {ok:true}; }

function getState(t, known) {
  const info = sess(t);
  const moved = movedTo();
  if (moved) return {moved: moved};
  if (!info) return ERR_SESSION;
  if (known && Number(known) === ver()) return {unchanged:true};
  return {session: info, state: state(ss(), info)};
}

function isOpen(i, cs, sc) {
  return EVENTS.slice(0, i).every(e => e.c.every(cat => {
    const m = sc[e.id + '-' + cat] || {};
    return cs.length > 0 && cs.every(c => m[c.id]);
  }));
}

function saveScores(t, key, map) {
  const info = sess(t);
  if (!info) return ERR_SESSION;
  if (info.role !== 'admin' && info.role !== 'juez') return {error:'Tu usuario no puede registrar resultados.'};
  const parts = String(key).split('-'), i = EVENTS.findIndex(e => e.id === parts[0]), ev = EVENTS[i];
  if (!ev || ev.c.indexOf(parts[1]) < 0) return {error:'Prueba desconocida.'};
  return locked(() => {
    const s = ss(), cs = centers(s), sc = scores(s);
    const asg = assignments(s)[ev.id] || [];
    if (info.role === 'juez' && asg.length && asg.indexOf(info.user) < 0)
      return {error:'Esta disciplina está asignada a otro juez (' + asg.join(', ') + ').'};
    if (info.role !== 'admin' && yes(cfgAll(s).orden_estricto) && !isOpen(i, cs, sc))
      return {error:'Esta disciplina aún no está abierta: termina antes las anteriores.'};
    const n = cs.length, names = {};
    cs.forEach(x => names[x.id] = x.name);
    const now = new Date();
    const keep = rows(s, SH.PUNT).filter(r => str(r[0]) !== key).map(r => HEAD['Puntuaciones'].map((_, j) => r[j] == null ? '' : r[j]));
    const add = [];
    Object.keys(map || {}).forEach(cid => {
      const pos = Math.round(Number(map[cid]));
      if (names[cid] == null || !(pos >= 1)) return;
      const p = Math.min(n, pos);
      add.push([key, ev.n, CATN[parts[1]], cid, names[cid], p, n - p + 1, info.user, now]);
    });
    writeRows(s, SH.PUNT, keep.concat(add));
    rebuildRanking(s);
    log(s, info, 'Resultados', ev.n + ' ' + CATN[parts[1]] + ': ' + add.length + '/' + n + ' centros');
    return {ok:true, version: bump()};
  });
}

function saveCenters(t, list) {
  return asStaff(t, (s, info) => {
    if (info.role !== 'admin') {
      const ids = (list || []).map(c => String(c.id));
      if (centers(s).some(o => ids.indexOf(o.id) < 0)) return {error:'Solo el administrador puede eliminar centros.'};
    }
    const data = (list || []).map((c, k) => [
      String(c.id || ('c' + Date.now() + k)), String(c.name || 'Centro'), String(c.color || '#475569'),
      String(c.user || '').trim(), String(c.pass || '')
    ]);
    writeRows(s, SH.CENTROS, data);
    rebuildRanking(s);
    log(s, info, 'Centros actualizados', data.length + ' centros');
    return {ok:true, version: bump()};
  });
}

function saveJudges(t, list) {
  return asStaff(t, (s, info) => {
    if (info.role !== 'admin' && (list || []).length < judges(s).length) return {error:'Solo el administrador puede eliminar jueces.'};
    const data = (list || []).map(j => [String(j.user || '').trim(), String(j.pass || '')]);
    writeRows(s, SH.JUECES, data);
    log(s, info, 'Jueces actualizados', data.length + ' jueces');
    return {ok:true, version: bump()};
  });
}

function saveOrgs(t, list) {
  return asAdmin(t, (s, info) => {
    const data = (list || []).map(o => [String(o.user || '').trim(), String(o.pass || '')]);
    writeRows(s, SH.ORG, data);
    log(s, info, 'Usuarios de organización actualizados', data.length + ' usuarios');
    return {ok:true, version: bump()};
  });
}

function saveAssignments(t, map) {
  return asStaff(t, (s, info) => {
    const valid = judges(s).map(j => j.user), data = [];
    EVENTS.forEach(e => {
      const us = ((map || {})[e.id] || []).map(u => String(u).trim());
      us.filter((u, k) => valid.indexOf(u) >= 0 && us.indexOf(u) === k).forEach(u => data.push([e.id, e.n, u]));
    });
    writeRows(s, SH.ASIG, data);
    log(s, info, 'Asignación de jueces', data.map(r => r[1] + ': ' + r[2]).join(' · ') || 'Sin asignaciones');
    return {ok:true, version: bump()};
  });
}

function saveCalendar(t, list) {
  return asStaff(t, (s, info) => {
    if (info.role !== 'admin') {
      const old = calendar(s), nw = list || [];
      const keys = nw.filter(r => r.id).map(r => r.id + '-' + r.cat);
      if (nw.length < old.length || old.some(r => r.id && keys.indexOf(r.id + '-' + r.cat) < 0))
        return {error:'Solo el administrador puede quitar actividades del calendario.'};
    }
    const data = (list || []).map(r => [r.id || '', r.name || '', r.cat || '', r.date || '', r.start || '', r.end || '', r.place || ''].map(String));
    writeRows(s, SH.CAL, data);
    log(s, info, 'Calendario actualizado', data.length + ' filas');
    return {ok:true, version: bump()};
  });
}

function saveSettings(t, p) {
  return asAdmin(t, (s, info) => {
    p = p || {};
    const patch = {};
    if ('logo' in p) {
      const l = String(p.logo || '');
      if (l.length > 49000) return {error:'El logotipo es demasiado pesado para guardarlo en la hoja.'};
      patch.logo = l;
    }
    if ('strictOrder' in p) patch.orden_estricto = p.strictOrder ? 'sí' : 'no';
    if ('centersReady' in p) patch.centros_listos = p.centersReady ? 'sí' : 'no';
    if ('movedTo' in p) {
      const u = String(p.movedTo || '').trim();
      if (u && !URL_RE.test(u)) return {error:'La nueva URL no es una URL /exec de Apps Script.'};
      if (u) PropertiesService.getScriptProperties().setProperty('MOVED_TO', u);
      else PropertiesService.getScriptProperties().deleteProperty('MOVED_TO');
      log(s, info, 'Servidor de datos', u ? 'Datos trasladados a ' + u : 'Traslado anulado');
    }
    if ('adminUser' in p) {
      if (!String(p.adminUser).trim()) return {error:'El usuario del administrador no puede quedar vacío.'};
      patch.admin_usuario = String(p.adminUser).trim();
    }
    if ('adminPass' in p) {
      if (String(p.adminPass).length < 4) return {error:'La contraseña del administrador debe tener al menos 4 caracteres.'};
      patch.admin_contraseña = String(p.adminPass);
    }
    if (Object.keys(patch).length) cfgSet(s, patch);
    if (Object.keys(patch).length) log(s, info, 'Ajustes', Object.keys(patch).filter(k => k !== 'admin_contraseña').join(', ') + ('admin_contraseña' in patch ? ' (contraseña cambiada)' : ''));
    return {ok:true, version: bump()};
  });
}

function resetScores(t) {
  return asAdmin(t, (s, info) => {
    writeRows(s, SH.PUNT, []);
    rebuildRanking(s);
    log(s, info, 'Puntuaciones reiniciadas', '');
    return {ok:true, version: bump()};
  });
}

/* ---------- Pestaña "Clasificación" (se regenera sola) ---------- */
function rebuildRanking(s) {
  const cs = centers(s), sc = scores(s), n = cs.length;
  const keys = [];
  EVENTS.forEach(e => e.c.forEach(c => keys.push({k: e.id + '-' + c, l: e.n + ' ' + (c === 'Mixto' ? 'Mixto' : c)})));
  const table = cs.map(c => {
    const per = keys.map(x => { const pos = (sc[x.k] || {})[c.id]; return pos ? n - pos + 1 : ''; });
    return {name: c.name, per: per, total: per.reduce((a, b) => a + (Number(b) || 0), 0)};
  }).sort((a, b) => b.total - a.total);
  const out = [['Posición','Centro','Total'].concat(keys.map(x => x.l))];
  table.forEach((r, i) => out.push([i + 1, r.name, r.total].concat(r.per)));
  const sh = s.getSheetByName(SH.CLAS);
  sh.clearContents();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
