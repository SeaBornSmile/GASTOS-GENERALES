/**
 * Mis gastos · backend en Google Apps Script
 *
 * Instalación (una sola vez):
 *  1. Abre tu Google Sheet > Extensiones > Apps Script y pega este archivo.
 *  2. Elige la función "setup" arriba y pulsa Ejecutar. Acepta los permisos.
 *     Crea las pestañas Gastos, Config y Resumen y genera tu clave secreta.
 *     La clave aparece en el registro de ejecución (y en Config!B2).
 *  3. Implementar > Nueva implementación > Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier usuario
 *     Copia la URL que termina en /exec.
 *  4. En la app, pega la URL y la clave.
 */

const HOJA_GASTOS = 'Gastos';
const HOJA_CONFIG = 'Config';
const HOJA_RESUMEN = 'Resumen';
const COLS = ['id', 'fecha', 'importe', 'categoria', 'concepto', 'pago', 'fijo_id'];
const COLS_FIJOS = ['id', 'nombre', 'importe', 'dia', 'categoria', 'pago'];
const FILA_FIJOS = 5; // cabecera de la tabla de fijos en Config

/* ---------- Instalación ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let g = ss.getSheetByName(HOJA_GASTOS);
  if (!g) {
    g = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() <= 1
      ? ss.getSheets()[0].setName(HOJA_GASTOS)
      : ss.insertSheet(HOJA_GASTOS);
  }
  g.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#0f6b57').setFontColor('#ffffff');
  g.setFrozenRows(1);
  g.getRange('A:A').setNumberFormat('@');
  g.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  g.getRange('C:C').setNumberFormat('#,##0.00 "€"');
  g.getRange('E:G').setNumberFormat('@');
  g.setColumnWidth(5, 220);

  let c = ss.getSheetByName(HOJA_CONFIG) || ss.insertSheet(HOJA_CONFIG);
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('TOKEN', token);
  }
  if (c.getRange('A1').getValue() !== 'presupuesto') {
    c.getRange('A1:B1').setValues([['presupuesto', 0]]);
  }
  c.getRange('A2:B2').setValues([['clave_app', token]]);
  c.getRange('B1').setNumberFormat('#,##0.00 "€"');
  c.getRange(FILA_FIJOS - 1, 1).setValue('Gastos fijos').setFontWeight('bold');
  c.getRange(FILA_FIJOS, 1, 1, COLS_FIJOS.length).setValues([COLS_FIJOS]).setFontWeight('bold');
  c.getRange(FILA_FIJOS, 1, 200, COLS_FIJOS.length).setNumberFormat('@');
  c.getRange(FILA_FIJOS + 1, 3, 200, 2).setNumberFormat('0.##');
  c.getRange('A1:A2').setFontWeight('bold');

  let r = ss.getSheetByName(HOJA_RESUMEN) || ss.insertSheet(HOJA_RESUMEN);
  r.clear();
  r.getRange('A1').setValue('Gasto por mes y categoría').setFontWeight('bold').setFontSize(13);
  r.getRange('A3').setFormula(
    '=IFERROR(QUERY({ARRAYFORMULA(IF(Gastos!B2:B="",,TEXT(Gastos!B2:B,"yyyy-mm"))),Gastos!D2:D,Gastos!C2:C},' +
    '"select Col1, sum(Col3) where Col1 is not null group by Col1 pivot Col2 order by Col1 desc label Col1 \'Mes\'",0),"Sin gastos todavía")'
  );
  r.getRange('A3:Z100').setNumberFormat('#,##0.00 "€"');
  r.getRange('A:A').setNumberFormat('@');
  r.setFrozenRows(3);

  Logger.log('Clave para la app: ' + token);
  return token;
}

/* ---------- HTTP ---------- */
function doGet(e) {
  return run_(() => {
    checkToken_(e.parameter.token);
    const action = e.parameter.action || 'list';
    if (action === 'list') return listAll_();
    if (action === 'ping') return { ok: true };
    throw new Error('Acción no válida: ' + action);
  });
}

function doPost(e) {
  return run_(() => {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    checkToken_(body.token);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ops = Array.isArray(body.ops) ? body.ops : [body];
      ops.forEach(applyOp_);
      SpreadsheetApp.flush();
      return listAll_();
    } finally {
      lock.releaseLock();
    }
  });
}

function run_(fn) {
  let out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(t) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token) throw new Error('Falta ejecutar setup() en Apps Script.');
  if (!t || t !== token) throw new Error('Clave incorrecta.');
}

/* ---------- Operaciones ---------- */
function applyOp_(op) {
  switch (op.action) {
    case 'add': (op.items || [op.item]).forEach(upsert_); break;
    case 'update': upsert_(op.item); break;
    case 'delete': deleteRow_(op.id); break;
    case 'config': saveConfig_(op.config || {}); break;
    default: throw new Error('Acción no válida: ' + op.action);
  }
}

function sheet_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('No existe la pestaña ' + name + '. Ejecuta setup() en Apps Script.');
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

function upsert_(it) {
  if (!it || !it.id) throw new Error('Gasto sin id');
  const importe = Number(it.importe);
  if (!(importe > 0)) throw new Error('Importe no válido');
  const sh = sheet_(HOJA_GASTOS);
  const row = [String(it.id), toDate_(it.fecha), importe, String(it.cat || 'Otros'), String(it.desc || ''), String(it.pago || ''), String(it.fijoId || '')];
  const r = findRow_(sh, it.id);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function deleteRow_(id) {
  const sh = sheet_(HOJA_GASTOS);
  const r = findRow_(sh, id);
  if (r < 0) return;
  // Sheets no deja borrar la última fila no congelada: en ese caso se vacía
  if (sh.getMaxRows() - sh.getFrozenRows() <= 1) sh.getRange(r, 1, 1, COLS.length).clearContent();
  else sh.deleteRow(r);
}

function saveConfig_(cfg) {
  const sh = sheet_(HOJA_CONFIG);
  if (cfg.presupuesto !== undefined) sh.getRange('B1').setValue(Number(cfg.presupuesto) || 0);
  if (Array.isArray(cfg.fijos)) {
    const last = sh.getLastRow();
    if (last > FILA_FIJOS) sh.getRange(FILA_FIJOS + 1, 1, last - FILA_FIJOS, COLS_FIJOS.length).clearContent();
    const rows = cfg.fijos.map(f => [String(f.id), String(f.nombre || ''), Number(f.importe) || 0, Number(f.dia) || 1, String(f.cat || 'Otros'), String(f.pago || '')]);
    if (rows.length) sh.getRange(FILA_FIJOS + 1, 1, rows.length, COLS_FIJOS.length).setValues(rows);
  }
}

function listAll_() {
  const tz = Session.getScriptTimeZone();
  const g = sheet_(HOJA_GASTOS);
  const n = g.getLastRow() - 1;
  const gastos = [];
  if (n > 0) {
    g.getRange(2, 1, n, COLS.length).getValues().forEach(v => {
      if (!v[0]) return;
      const fecha = v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]).slice(0, 10);
      gastos.push({ id: String(v[0]), fecha: fecha, importe: Number(v[2]) || 0, cat: String(v[3]), desc: String(v[4]), pago: String(v[5]), fijoId: String(v[6] || '') || undefined });
    });
  }
  const c = sheet_(HOJA_CONFIG);
  const fijos = [];
  const last = c.getLastRow();
  if (last > FILA_FIJOS) {
    c.getRange(FILA_FIJOS + 1, 1, last - FILA_FIJOS, COLS_FIJOS.length).getValues().forEach(v => {
      if (!v[0]) return;
      fijos.push({ id: String(v[0]), nombre: String(v[1]), importe: Number(v[2]) || 0, dia: Number(v[3]) || 1, cat: String(v[4]), pago: String(v[5]) });
    });
  }
  return { gastos: gastos, config: { presupuesto: Number(c.getRange('B1').getValue()) || 0, fijos: fijos } };
}
