/**
 * Mis gastos · backend en Google Apps Script · versión 2
 *
 * Instalación (una sola vez):
 *  1. En tu Google Sheet: Extensiones > Apps Script, pega este archivo y guarda.
 *  2. Elige la función "setup" y pulsa Ejecutar. Acepta los permisos.
 *     Crea o actualiza las pestañas y genera tu clave (Config!B2).
 *  3. Implementar > Nueva implementación > Aplicación web
 *     (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario). Copia la URL /exec.
 *
 * Actualizar desde la versión 1:
 *  1. Sustituye todo el código por este y guarda.
 *  2. Ejecuta "setup" (migra tus datos: no se pierde nada).
 *  3. Implementar > Gestionar implementaciones > (lápiz) Editar > Versión: Nueva versión > Implementar.
 *     Así la URL /exec no cambia.
 */

const VERSION = 2;
const H_MOV = 'Movimientos';
const H_CONFIG = 'Config';
const H_FIJOS = 'Fijos';
const H_CATS = 'Categorias';
const H_OBJ = 'Objetivos';
const H_RESUMEN = 'Resumen';

const COLS_MOV = ['id', 'fecha', 'importe', 'categoria', 'concepto', 'pago', 'fijo_id', 'tipo', 'persona', 'notas', 'subcategoria'];
const COLS_FIJOS = ['id', 'nombre', 'importe', 'dia', 'categoria', 'pago', 'tipo', 'persona'];
const COLS_CATS = ['id', 'emoji', 'nombre', 'tipo', 'padre', 'limite'];
const COLS_OBJ = ['id', 'emoji', 'nombre', 'meta', 'ahorrado', 'fecha_limite'];
const CONFIG_KEYS = ['presupuesto', 'clave_app', 'personas', 'saldo_inicial', 'ahorro_mensual'];

/* ================= Instalación y migración ================= */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // Movimientos (antes "Gastos")
  let m = ss.getSheetByName(H_MOV);
  if (!m) {
    const old = ss.getSheetByName('Gastos');
    if (old) m = old.setName(H_MOV);
    else if (ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() <= 1) m = ss.getSheets()[0].setName(H_MOV);
    else m = ss.insertSheet(H_MOV);
  }
  header_(m, COLS_MOV);
  m.getRange('A:A').setNumberFormat('@');
  m.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  m.getRange('C:C').setNumberFormat('#,##0.00 "€"');
  m.getRange('D:K').setNumberFormat('@');
  m.setColumnWidth(5, 220);

  // Config: clave/valor en A1:B5. Migra la tabla de fijos de la versión 1 (A5:F)
  const c = ss.getSheetByName(H_CONFIG) || ss.insertSheet(H_CONFIG);
  const fijosV1 = [];
  if (String(c.getRange('A5').getValue()) === 'id' && String(c.getRange('B5').getValue()) === 'nombre') {
    const last = c.getLastRow();
    if (last > 5) c.getRange(6, 1, last - 5, 6).getValues().forEach(v => { if (v[0]) fijosV1.push(v.concat(['gasto', ''])); });
    c.getRange(4, 1, Math.max(1, c.getLastRow() - 3), 6).clearContent().setFontWeight('normal');
  }
  const prev = {};
  c.getRange(1, 1, 10, 2).getValues().forEach(r => { if (r[0]) prev[String(r[0])] = r[1]; });
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) { token = Utilities.getUuid().replace(/-/g, ''); props.setProperty('TOKEN', token); }
  const defaults = { presupuesto: 0, clave_app: token, personas: '', saldo_inicial: 0, ahorro_mensual: 0 };
  c.getRange(1, 1, CONFIG_KEYS.length, 2).setValues(CONFIG_KEYS.map(k => [k, k === 'clave_app' ? token : (prev[k] !== undefined && prev[k] !== '' ? prev[k] : defaults[k])]));
  c.getRange('A1:A5').setFontWeight('bold');
  c.getRange('B1').setNumberFormat('#,##0.00 "€"');
  c.getRange('B4:B5').setNumberFormat('#,##0.00 "€"');
  c.getRange('B3').setNumberFormat('@');
  c.getRange('D1').setValue('personas: nombres separados por comas (ej.: Borja, Ana)');
  c.getRange('D1').setFontColor('#7d868f');

  // Tablas
  const f = ss.getSheetByName(H_FIJOS) || ss.insertSheet(H_FIJOS);
  header_(f, COLS_FIJOS);
  if (fijosV1.length && f.getLastRow() <= 1) f.getRange(2, 1, fijosV1.length, COLS_FIJOS.length).setValues(fijosV1);
  header_(ss.getSheetByName(H_CATS) || ss.insertSheet(H_CATS), COLS_CATS);
  header_(ss.getSheetByName(H_OBJ) || ss.insertSheet(H_OBJ), COLS_OBJ);

  // Resumen: gasto por mes y categoría (solo gastos)
  const r = ss.getSheetByName(H_RESUMEN) || ss.insertSheet(H_RESUMEN);
  r.clear();
  r.getRange('A1').setValue('Gasto por mes y categoría').setFontWeight('bold').setFontSize(13);
  r.getRange('A3').setFormula(
    '=IFERROR(QUERY({ARRAYFORMULA(IF(Movimientos!B2:B="",,TEXT(Movimientos!B2:B,"yyyy-mm"))),Movimientos!D2:D,Movimientos!C2:C,' +
    'ARRAYFORMULA(IF(Movimientos!H2:H="","gasto",Movimientos!H2:H))},' +
    '"select Col1, sum(Col3) where Col1 is not null and Col4 <> \'ingreso\' group by Col1 pivot Col2 order by Col1 desc label Col1 \'Mes\'",0),"Sin gastos todavía")'
  );
  r.getRange('B3:Z200').setNumberFormat('#,##0.00 "€"');
  r.setFrozenRows(3);

  ss.setActiveSheet(m);
  Logger.log('Listo (versión ' + VERSION + '). Clave para la app: ' + token);
  return token;
}

function header_(sh, cols) {
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#0f6b57').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

/* ================= HTTP ================= */
function doGet(e) {
  return run_(() => {
    checkToken_(e.parameter.token);
    if ((e.parameter.action || 'list') === 'ping') return { ok: true, version: VERSION };
    return listAll_();
  });
}

function doPost(e) {
  return run_(() => {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    checkToken_(body.token);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      (Array.isArray(body.ops) ? body.ops : [body]).forEach(applyOp_);
      SpreadsheetApp.flush();
      return listAll_();
    } finally { lock.releaseLock(); }
  });
}

function run_(fn) {
  let out;
  try { out = { ok: true, data: fn() }; }
  catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(t) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token) throw new Error('Falta ejecutar setup() en Apps Script.');
  if (!t || t !== token) throw new Error('Clave incorrecta.');
}

/* ================= Operaciones ================= */
function applyOp_(op) {
  switch (op.action) {
    case 'add': (op.items || [op.item]).forEach(upsertMov_); break;
    case 'update': upsertMov_(op.item); break;
    case 'delete': deleteMov_(op.id); break;
    case 'config': saveConfig_(op.config || {}); break;
    case 'cats': writeTable_(H_CATS, COLS_CATS, (op.cats || []).map(x => [x.id, x.emoji || '', x.nombre, x.tipo || 'gasto', x.padre || '', Number(x.limite) || ''])); break;
    case 'objetivos': writeTable_(H_OBJ, COLS_OBJ, (op.objetivos || []).map(x => [x.id, x.emoji || '', x.nombre, Number(x.meta) || 0, Number(x.ahorrado) || 0, x.fecha || ''])); break;
    case 'renameCat': renameCat_(op.from, op.to, op.col || 'categoria'); break;
    default: throw new Error('Acción no válida: ' + op.action);
  }
}

function sheet_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('Falta la pestaña ' + name + '. Ejecuta setup() en Apps Script.');
  return s;
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
  for (let i = 0; i < n; i++) if (ids[i][0] === String(id)) return i + 2;
  return -1;
}

function toDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Fecha no válida: ' + s);
  return new Date(+m[1], +m[2] - 1, +m[3], 12);
}

function upsertMov_(it) {
  if (!it || !it.id) throw new Error('Movimiento sin id');
  const importe = Number(it.importe);
  if (!(importe > 0)) throw new Error('Importe no válido');
  const sh = sheet_(H_MOV);
  const row = [String(it.id), toDate_(it.fecha), importe, String(it.cat || 'Otros'), String(it.desc || ''), String(it.pago || ''),
    String(it.fijoId || ''), it.tipo === 'ingreso' ? 'ingreso' : 'gasto', String(it.persona || ''), String(it.notas || ''), String(it.sub || '')];
  const r = findRow_(sh, it.id);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function deleteMov_(id) {
  const sh = sheet_(H_MOV);
  const r = findRow_(sh, id);
  if (r < 0) return;
  if (sh.getMaxRows() - sh.getFrozenRows() <= 1) sh.getRange(r, 1, 1, COLS_MOV.length).clearContent();
  else sh.deleteRow(r);
}

function renameCat_(from, to, col) {
  if (!from || !to || from === to) return;
  const sh = sheet_(H_MOV);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ci = col === 'subcategoria' ? 11 : 4;
  sh.getRange(2, ci, n, 1).createTextFinder(String(from)).matchEntireCell(true).replaceAllWith(String(to));
}

function writeTable_(name, cols, rows) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
}

function readTable_(name, cols) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().filter(v => v[0] !== '').map(v => {
    const o = {}; cols.forEach((c, i) => o[c] = v[i]); return o;
  });
}

function saveConfig_(cfg) {
  const sh = sheet_(H_CONFIG);
  const set = (key, val) => { const i = CONFIG_KEYS.indexOf(key); if (i >= 0) sh.getRange(i + 1, 2).setValue(val); };
  if (cfg.presupuesto !== undefined) set('presupuesto', Number(cfg.presupuesto) || 0);
  if (cfg.saldoInicial !== undefined) set('saldo_inicial', Number(cfg.saldoInicial) || 0);
  if (cfg.ahorroMensual !== undefined) set('ahorro_mensual', Number(cfg.ahorroMensual) || 0);
  if (Array.isArray(cfg.personas)) set('personas', cfg.personas.join(', '));
  if (Array.isArray(cfg.fijos)) writeTable_(H_FIJOS, COLS_FIJOS, cfg.fijos.map(f => [String(f.id), f.nombre || '', Number(f.importe) || 0, Number(f.dia) || 1, f.cat || 'Otros', f.pago || '', f.tipo === 'ingreso' ? 'ingreso' : 'gasto', f.persona || '']));
}

function listAll_() {
  const tz = Session.getScriptTimeZone();
  const fmt = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').slice(0, 10);
  const movs = readTable_(H_MOV, COLS_MOV).map(o => ({
    id: String(o.id), fecha: fmt(o.fecha), importe: Number(o.importe) || 0, cat: String(o.categoria), desc: String(o.concepto), pago: String(o.pago),
    fijoId: String(o.fijo_id || '') || undefined, tipo: o.tipo === 'ingreso' ? 'ingreso' : 'gasto', persona: String(o.persona || ''), notas: String(o.notas || ''), sub: String(o.subcategoria || '')
  }));
  const c = sheet_(H_CONFIG).getRange(1, 1, CONFIG_KEYS.length, 2).getValues();
  const val = k => { const r = c.find(x => x[0] === k); return r ? r[1] : ''; };
  return {
    version: VERSION,
    gastos: movs,
    config: {
      presupuesto: Number(val('presupuesto')) || 0,
      saldoInicial: Number(val('saldo_inicial')) || 0,
      ahorroMensual: Number(val('ahorro_mensual')) || 0,
      personas: String(val('personas') || '').split(',').map(s => s.trim()).filter(Boolean),
      fijos: readTable_(H_FIJOS, COLS_FIJOS).map(f => ({ id: String(f.id), nombre: String(f.nombre), importe: Number(f.importe) || 0, dia: Number(f.dia) || 1, cat: String(f.categoria), pago: String(f.pago), tipo: f.tipo === 'ingreso' ? 'ingreso' : 'gasto', persona: String(f.persona || '') }))
    },
    cats: readTable_(H_CATS, COLS_CATS).map(x => ({ id: String(x.id), emoji: String(x.emoji), nombre: String(x.nombre), tipo: x.tipo === 'ingreso' ? 'ingreso' : 'gasto', padre: String(x.padre || ''), limite: Number(x.limite) || 0 })),
    objetivos: readTable_(H_OBJ, COLS_OBJ).map(x => ({ id: String(x.id), emoji: String(x.emoji), nombre: String(x.nombre), meta: Number(x.meta) || 0, ahorrado: Number(x.ahorrado) || 0, fecha: fmt(x.fecha_limite) }))
  };
}
