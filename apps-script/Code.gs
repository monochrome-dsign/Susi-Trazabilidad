/**
 * SUSI — Trazabilidad. Backend sobre Google Sheets.
 * Se pega en: hoja de cálculo > Extensiones > Apps Script.
 * La clave de acceso se genera al ejecutar setup() y NO se guarda en ningún archivo.
 */
const TABS = {
  MateriaPrima: ['id','lote','producto','proveedor','cantidad','unidad','entrada','caducidad','camara'],
  Produccion: ['id','lote','producto','fecha','turno','operario','cantidad','unidad'],
  Consumos: ['prId','mpId','lote','cantidad'],
  ProductosTerminados: ['id','prId','barcode','nombre','formato','cantidad','unidad','envasado','caducidad','almacen'],
  Expediciones: ['id','ptId','cliente','fecha','cantidad','unidad','destino']
};

// Ejecutar UNA vez desde el editor: crea pestañas que falten y genera la clave (ver Registro de ejecución).
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) s.getRange(1, 1, 1, TABS[n].length).setValues([TABS[n]]).setFontWeight('bold');
  });
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('CLAVE')) p.setProperty('CLAVE', Utilities.getUuid().slice(0, 8).toUpperCase());
  Logger.log('CLAVE DE ACCESO: ' + p.getProperty('CLAVE'));
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function autorizado(k) { return !!k && k === PropertiesService.getScriptProperties().getProperty('CLAVE'); }
// ID único (aleatorio): Date.now() puede repetirse si dos registros caen en el mismo milisegundo
function uid(pref) { return pref + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function hoy() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function doGet(e) {
  try { return autorizado(e.parameter.key) ? out(getAll()) : out({ error: 'Clave incorrecta' }); }
  catch (err) { return out({ error: String(err.message || err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita que dos usuarios escriban a la vez y rompan el stock
  try {
    const r = JSON.parse(e.postData.contents);
    if (!autorizado(r.key)) return out({ error: 'Clave incorrecta' });
    const fn = { addMateria: addMateria, addProduccion: addProduccion, addExpedicion: addExpedicion }[r.action];
    if (!fn) return out({ error: 'Acción desconocida' });
    fn(r.data);
    return out(getAll());
  } catch (err) { return out({ error: String(err.message || err) }); }
  finally { lock.releaseLock(); }
}

function read(name) {
  const v = SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (k === 'cantidad') x = Number(x);
      if (k === 'barcode') x = String(x);
      o[k] = x;
    });
    return o;
  });
}

function append(name, o) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  const row = TABS[name].map(k => o[k] === undefined ? '' : o[k]);
  s.getRange(s.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

function getAll() {
  const cons = read('Consumos');
  return {
    materias: read('MateriaPrima'),
    producciones: read('Produccion').map(p => Object.assign(p, { consumos: cons.filter(c => c.prId === p.id) })),
    productos: read('ProductosTerminados'),
    expediciones: read('Expediciones')
  };
}

// MP-AAAA-MMDD-NNN / PR-AAAA-MMDD-NNN (secuencia por día)
function lote(pref, tab, fecha) {
  const d = fecha.split('-'), base = pref + '-' + d[0] + '-' + d[1] + d[2] + '-';
  const n = read(tab).filter(r => String(r.lote).indexOf(base) === 0).length + 1;
  return base + String(n).padStart(3, '0');
}

// Código EAN-13 de uso interno (prefijo 29 = circulación restringida): 29 + AAMMDD + secuencia + dígito de control
function barcode(fecha) {
  const d = fecha.split('-'), pre = '29' + d[0].slice(2) + d[1] + d[2];
  const n = read('ProductosTerminados').filter(p => String(p.barcode).indexOf(pre) === 0).length + 1;
  const b = pre + String(n).padStart(4, '0');
  let s = 0;
  for (let i = 0; i < 12; i++) s += Number(b[i]) * (i % 2 ? 3 : 1);
  return b + ((10 - s % 10) % 10);
}

function addMateria(d) {
  const cant = Number(d.cantidad);
  if (!d.producto || !d.proveedor || !d.caducidad || !(cant > 0)) throw new Error('Faltan datos de la entrada');
  const entrada = d.entrada || hoy();
  append('MateriaPrima', {
    id: uid('mp'), lote: lote('MP', 'MateriaPrima', entrada), producto: d.producto,
    proveedor: d.proveedor, cantidad: cant, unidad: 'kg', entrada: entrada, caducidad: d.caducidad,
    camara: d.camara || 'Sin asignar'
  });
}

function addProduccion(d) {
  ['producto', 'operario', 'formato', 'caducidad'].forEach(k => { if (!d[k]) throw new Error('Falta el campo ' + k); });
  const cant = Number(d.cantidad), uds = Number(d.unidades);
  if (!(cant > 0) || !(uds > 0)) throw new Error('Cantidades no válidas');
  const fecha = d.fecha || hoy();
  const mps = read('MateriaPrima');
  // 1) validar todo y calcular el stock restante ANTES de escribir nada
  const usos = (d.consumos || []).map(c => {
    const mp = mps.find(m => m.id === c.mpId), q = Number(c.cantidad);
    if (!mp || !(q > 0)) throw new Error('Materia prima no válida');
    if (q > mp.cantidad) throw new Error('Stock insuficiente de ' + mp.producto + ' (' + mp.lote + '): quedan ' + mp.cantidad + ' kg');
    mp.cantidad = +(mp.cantidad - q).toFixed(2);
    return { mp: mp, q: q };
  });
  if (!usos.length) throw new Error('Indica al menos una materia prima');
  // 2) descontar stock
  const s = SpreadsheetApp.getActive().getSheetByName('MateriaPrima');
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]);
  usos.forEach(u => s.getRange(ids.indexOf(u.mp.id) + 1, 5).setValue(u.mp.cantidad));
  // 3) registrar lote, consumos y producto terminado
  const id = uid('pr');
  append('Produccion', { id: id, lote: lote('PR', 'Produccion', fecha), producto: d.producto, fecha: fecha, turno: d.turno, operario: d.operario, cantidad: cant, unidad: 'kg' });
  usos.forEach(u => append('Consumos', { prId: id, mpId: u.mp.id, lote: u.mp.lote, cantidad: u.q }));
  append('ProductosTerminados', { id: uid('pt'), prId: id, barcode: barcode(fecha), nombre: d.producto, formato: d.formato, cantidad: uds, unidad: 'bandejas', envasado: fecha, caducidad: d.caducidad, almacen: d.almacen || 'Sin asignar' });
}

// Salida de producto terminado hacia un cliente: valida y descuenta el stock del lote
function addExpedicion(d) {
  const cant = Number(d.cantidad);
  if (!d.ptId || !d.cliente || !(cant > 0)) throw new Error('Faltan datos de la expedición');
  const p = read('ProductosTerminados').find(x => x.id === d.ptId);
  if (!p) throw new Error('Producto no encontrado');
  if (cant > p.cantidad) throw new Error('Stock insuficiente de ' + p.nombre + ': quedan ' + p.cantidad + ' ' + p.unidad);
  const s = SpreadsheetApp.getActive().getSheetByName('ProductosTerminados');
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]);
  s.getRange(ids.indexOf(p.id) + 1, 6).setValue(+(p.cantidad - cant).toFixed(2));
  append('Expediciones', { id: uid('ex'), ptId: p.id, cliente: d.cliente, fecha: d.fecha || hoy(), cantidad: cant, unidad: p.unidad, destino: d.destino || '' });
}
