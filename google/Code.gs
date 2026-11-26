// Registro de Dosis Hidroponia: usuarios, mediciones de tanques y ajustes compartidos.
// Proyecto de Apps Script publicado como "Aplicación web".

// Planilla "Hidro Dosis Usuarios"
const SS_ID = '1vxmWFBLpxiWxihmsyGTzwQ2NOoUXO5o54gotwPuVGYA';
let ss__ = null;
function ss_() { return ss__ || (ss__ = SpreadsheetApp.openById(SS_ID)); }
const LOG_HEADERS = ['Fecha', 'Tanque', 'Usuario', 'Litros', 'ppm', 'Objetivo', 'Etapa', 'Agua', 'Indicación', 'Estado', 'pH', 'Nota'];
const COLS = LOG_HEADERS.length;
const TANK_NAMES = { T1: 'Tanque 1', T2: 'Tanque 2', T3: 'Tanque 3', T4: 'Tanque 4', T5: 'Tanque 5', T6: 'Tanque 6', T7: 'Tanque 7', T8: 'Tanque 8',
  P1: 'Plantinero 1', P2: 'Plantinero 2', P3: 'Plantinero 3', P4: 'Plantinero 4' };

function sheet_(name, headers) {
  let s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    s.appendRow(headers);
    s.setFrozenRows(1);
  }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props_() {
  return PropertiesService.getScriptProperties();
}

function getJSON_(key, def) {
  const v = props_().getProperty(key);
  return v ? JSON.parse(v) : def;
}

function setJSON_(key, value) {
  props_().setProperty(key, JSON.stringify(value));
}

function hash_(pin) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'hidro:' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}

// La primera vez que se usa la contraseña de ajustes, queda guardada acá.
// Si cambiás la contraseña de ajustes: Configuración del proyecto > Propiedades del script > borrar ADMIN.
function checkPin_(pin) {
  if (!pin) return false;
  const h = hash_(String(pin));
  const saved = props_().getProperty('ADMIN');
  if (!saved) {
    props_().setProperty('ADMIN', h);
    return true;
  }
  return saved === h;
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function str_(v, max) {
  return String(v == null ? '' : v).slice(0, max || 80);
}

// Usuarios: se lee solo cuando hace falta
function users_() {
  const sh = sheet_('Usuarios', ['ID', 'Nombre', 'Dispositivo', 'Primer ingreso', 'Último ingreso', 'Ingresos']);
  const data = sh.getDataRange().getValues();
  const rowOf = (id) => {
    for (let i = 1; i < data.length; i++) if (data[i][0] === id) return i + 1;
    return 0;
  };
  const nameOf = (id) => {
    const r = rowOf(String(id || ''));
    return r ? String(data[r - 1][1]) : 'Sin registrar';
  };
  return { sh: sh, data: data, rowOf: rowOf, nameOf: nameOf };
}

// Deja una medición guardada con los campos que corresponden (para "deshacer")
function entry_(p) {
  return {
    time: str_(p.time, 30), user: str_(p.user, 40), liters: num_(p.liters), ppm: num_(p.ppm), target: num_(p.target),
    stage: str_(p.stage, 20), fill: !!p.fill, kind: str_(p.kind, 10), summary: str_(p.summary, 300),
    ph: p.ph ? num_(p.ph) : '', note: str_(p.note, 120),
    done: p.done && p.done.time ? { time: str_(p.done.time, 30), user: str_(p.done.user, 40) } : null,
  };
}

// Hoja de mediciones. Si es la hoja vieja (sin pH ni Nota), le agrega los dos títulos.
function logSheet_() {
  const s = sheet_('Mediciones', LOG_HEADERS);
  if (!s.getRange(1, COLS).getValue()) s.getRange(1, 1, 1, COLS).setValues([LOG_HEADERS]);
  return s;
}

// Dispositivos bloqueados por el dueño. Se guarda aparte de la hoja Usuarios,
// así un dispositivo sigue bloqueado aunque se lo borre de la lista.
function blockedIds_() {
  return getJSON_('BLOCKED', []);
}

function isBlocked_(id) {
  return !!id && blockedIds_().indexOf(String(id)) >= 0;
}

const ADMIN_ACTIONS = ['fin', 'finset', 'stock', 'stockadd', 'stockdel', 'orderadd', 'orderset', 'sup', 'supset', 'fumadd', 'fumdel', 'history', 'setconfig', 'list', 'rename', 'block', 'remove', 'people', 'setpeople', 'unlink', 'hoursadmin', 'hourspay', 'hoursunpay', 'paydel', 'hoursedit', 'empentries', 'empwipe'];

// ---------- Horas de trabajo ----------
// Hoja "Horas": una fila por carga. Estado "Pendiente" hasta que el dueño la paga; ahí pasa a "Pagado".
// Hoja "Pagos": una fila por pago (con el total de horas y de plata).
// Empleados: las personas con acceso a Horas (ver "Personas y accesos" más abajo).
const HOURS_HEADERS = ['ID', 'Cargado', 'Día', 'Empleado', 'Horas', 'Trabajo', 'Estado', 'Pagado el', 'Pago', 'Dispositivo', 'Cargado por'];
const PAY_HEADERS = ['ID', 'Fecha', 'Empleado', 'Horas', 'Valor hora', 'Total', 'Desde', 'Hasta', 'Cargas', 'Pagado por', 'Estado'];
const HARV_HEADERS = ['ID', 'Fecha', 'Tanques', 'Paquetes', 'Nota', 'Cargado por', 'Dispositivo', 'Cargado el', 'Cajones', 'Detalle cajones', 'Agujeros'];
const HARV_BOX = 15; // paquetes en un cajón común

// ---------- Finanzas (solo el dueño) ----------
// Clientes: una fila por precio (si cambia el precio, se agrega otra fila con la fecha desde la que vale).
// Ventas: el reparto de cada cosecha entre clientes (se reescribe entero cada vez que se guarda esa cosecha).
// Gastos del mes: luz, monotributo y otros. Vehículo: combustible (con los km del tablero), service, seguro, patente, cubiertas.
const CLI_HEADERS = ['ID', 'Cliente', 'Precio por paquete', 'Desde', 'Cargado el'];
const SALE_HEADERS = ['ID', 'Cosecha', 'Fecha cosecha', 'Cliente ID', 'Cliente', 'Cajones', 'Paquetes', 'Precio por paquete', 'Total', 'Cargado el'];
const FCOST_HEADERS = ['ID', 'Mes', 'Tipo', 'Detalle', 'Monto', 'Cargado el'];
const VEH_HEADERS = ['ID', 'Fecha', 'Tipo', 'Monto', 'Litros', 'Km', 'Detalle', 'Cargado el'];
const VEH_TYPES = ['Combustible', 'Service', 'Seguro', 'Patente', 'Cubiertas'];
function finData_() {
  return {
    ok: true,
    clients: rows_('Clientes', CLI_HEADERS.length).map((r) => ({ id: String(r[0]), name: String(r[1] || ''), price: num_(r[2]), since: ymd_(r[3]) })).filter((x) => x.id && x.name),
    sales: rows_('Ventas', SALE_HEADERS.length).slice(-6000).map((r) => ({ id: String(r[0]), harvest: String(r[1]), day: ymd_(r[2]), client: String(r[3]), name: String(r[4] || ''), boxes: num_(r[5]), packs: num_(r[6]), price: num_(r[7]), total: num_(r[8]) })).filter((x) => x.id && x.harvest),
    costs: rows_('Gastos del mes', FCOST_HEADERS.length).map((r) => ({ id: String(r[0]), month: String(r[1]).slice(0, 7), type: String(r[2] || ''), detail: String(r[3] || ''), amount: num_(r[4]) })).filter((x) => x.id && /^\d{4}-\d{2}$/.test(x.month)),
    veh: rows_('Vehículo', VEH_HEADERS.length).map((r) => ({ id: String(r[0]), day: ymd_(r[1]), type: String(r[2] || ''), amount: num_(r[3]), liters: num_(r[4]), km: num_(r[5]), detail: String(r[6] || '') })).filter((x) => x.id && x.day),
  };
}
function addRows_(name, headers, rows) {
  if (!rows.length) return;
  const sh = sheet_(name, headers);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}
const okDay_ = (d) => /^\d{4}-\d{2}-\d{2}$/.test(String(d));
const FUM_HEADERS = ['ID', 'Fumigación', 'Fecha', 'Tanques', 'Producto', 'Tipo', 'Para qué', 'Dosis', 'Carencia (días)', 'Se puede cosechar desde', 'Cargado por', 'Cargado el'];

// ---------- Proveedores, productos y precios (solo el dueño) ----------
// Proveedores y productos se reescriben enteros en cada guardado (son pocos).
// Los precios se agregan: cada fila es el precio que se pagó un día, así queda el historial de aumentos.
// La receta de costos (con qué se prepara cada tanque A, B y C) va en una propiedad.
const SUP_HEADERS = ['ID', 'Proveedor', 'Teléfono', 'Notas'];
const SPROD_HEADERS = ['ID', 'Proveedor ID', 'Proveedor', 'Producto', 'Cantidad', 'Unidad', 'Marca'];
const PRICE_HEADERS = ['ID', 'Producto ID', 'Producto', 'Proveedor', 'Fecha', 'Precio'];
const PROD_UNITS = ['L', 'ml', 'kg', 'g', 'unidad'];

function supData_() {
  return {
    ok: true,
    suppliers: rows_('Proveedores', SUP_HEADERS.length).map((r) => ({ id: String(r[0]), name: String(r[1] || ''), phone: String(r[2] || ''), notes: String(r[3] || '') })),
    products: rows_('Productos de proveedores', SPROD_HEADERS.length).map((r) => ({ id: String(r[0]), sup: String(r[1] || ''), name: String(r[3] || ''), size: num_(r[4]), unit: String(r[5] || ''), brand: String(r[6] || '') })),
    prices: rows_('Precios', PRICE_HEADERS.length).slice(-5000).map((r) => ({ id: String(r[0]), product: String(r[1] || ''), day: ymd_(r[4]), price: num_(r[5]) })),
    recipe: getJSON_('COSTRECIPE', null),
  };
}

function writeAll_(name, headers, rows) {
  const sh = sheet_(name, headers);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]); // por si se agregó una columna nueva (ej.: Marca)
  const n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ---------- Stock de insumos (solo el dueño) ----------
// Hoja "Stock": un movimiento por fila. Tipos: Conteo (deja el stock en esa cantidad), Compra (+),
// Preparación (− químicos sólidos y + litros de Tanque A/B/C), Ajuste (+ o −).
// Lo que se usa de los tanques A, B y C sale solo de las correcciones confirmadas en "Mediciones".
const STOCK_HEADERS = ['ID', 'Fecha', 'Insumo', 'Cantidad', 'Unidad', 'Tipo', 'Detalle', 'Proveedor', 'Precio por unidad', 'Total', 'Cargado el'];
const STOCK_TYPES = ['Conteo', 'Compra', 'Preparación', 'Ajuste'];
const STOCK_UNITS = ['kg', 'L', 'unidad'];

// "4,77" o "1.234,5" -> número; ml -> L
function amt_(v, u) {
  const n = parseFloat(String(v).replace(/\./g, '').replace(',', '.'));
  return isFinite(n) ? (u === 'ml' ? n / 1000 : n) : 0;
}

// Litros de A, B y C de cada corrección confirmada desde "since" (las deshechas no cuentan)
function tankUse_(since) {
  const s = ss_().getSheetByName('Mediciones');
  if (!s || s.getLastRow() < 2 || !since) return [];
  const n = s.getLastRow(), start = Math.max(2, n - 19999);
  const vals = s.getRange(start, 1, n - start + 1, 10).getValues();
  const open = {}, out = [];
  vals.forEach((r) => {
    if (!(r[0] instanceof Date) || r[0] < since) return;
    const tank = String(r[1]), st = String(r[9]), sum = String(r[8] || '');
    if (st === 'Corregido') {
      const e = { time: r[0].toISOString(), tank: tank, A: 0, B: 0, C: 0 };
      ['A', 'B', 'C'].forEach((k) => {
        const m = sum.match(new RegExp('\\b' + k + ' ([\\d.,]+) (L|ml)'));
        if (m) e[k] = amt_(m[1], m[2]);
      });
      out.push(e);
      (open[tank] = open[tank] || []).push(e);
    } else if (st === 'Anulado' && /correcci/i.test(sum) && open[tank] && open[tank].length) {
      const e = open[tank].pop();
      e.A = e.B = e.C = 0;
    }
  });
  return out.filter((e) => e.A || e.B || e.C);
}

function stockData_() {
  const rows = rows_('Stock', STOCK_HEADERS.length);
  const ledger = rows.map((r) => ({
    id: String(r[0]), day: ymd_(r[1]), item: String(r[2] || ''), qty: num_(r[3]), unit: String(r[4] || ''), type: String(r[5] || ''),
    detail: String(r[6] || ''), sup: String(r[7] || ''), price: num_(r[8]), total: num_(r[9]), at: iso_(r[10]),
  }));
  let since = null;
  rows.forEach((r) => { if (/^Tanque [ABC]$/.test(String(r[2])) && r[1] instanceof Date && (!since || r[1] < since)) since = r[1]; });
  if (since) { since = new Date(since); since.setHours(0, 0, 0, 0); }
  return { ok: true, ledger: ledger, use: tankUse_(since), orders: ordersData_() };
}

// ---------- Pedidos a proveedores (solo el dueño) ----------
// Hoja "Pedidos": una fila por producto pedido. Los de un mismo pedido comparten el comienzo del ID (pedido~1, pedido~2...).
// Estado: Pedido (esperando que llegue), Recibido (se cargó la compra) o Cancelado.
const ORDER_HEADERS = ['ID', 'Fecha', 'Proveedor ID', 'Proveedor', 'Insumo', 'Producto ID', 'Cantidad', 'Presentación', 'Estado', 'Cargado el'];
const ORDER_STATES = ['Pedido', 'Recibido', 'Cancelado'];

function ordersData_() {
  return rows_('Pedidos', ORDER_HEADERS.length).slice(-500).map((r) => ({
    id: String(r[0]), day: ymd_(r[1]), sup: String(r[2] || ''), supName: String(r[3] || ''), item: String(r[4] || ''), product: String(r[5] || ''),
    n: num_(r[6]), detail: String(r[7] || ''), status: String(r[8] || ''), at: iso_(r[9]),
  }));
}

// Cambia el estado de todos los productos de un pedido (se escribe la columna de una vez)
function setOrderState_(group, state) {
  const sh = ss_().getSheetByName('Pedidos');
  if (!group || !sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1, ids = sh.getRange(2, 1, n, 1).getValues(), st = sh.getRange(2, 9, n, 1).getValues();
  let changed = false;
  ids.forEach((r, i) => { const id = String(r[0]); if (id === group || id.indexOf(group + '~') === 0) { st[i][0] = state; changed = true; } });
  if (changed) sh.getRange(2, 9, n, 1).setValues(st);
}

// ---------- Fumigaciones ----------
// Una fila por producto. Los productos aplicados juntos comparten la columna "Fumigación".
// "Tanques" va con los nombres separados por coma.
function fumRows_() {
  return rows_('Fumigaciones', FUM_HEADERS.length).slice(-600).map((r) => ({
    id: String(r[0]), group: String(r[1]), day: ymd_(r[2]), tanks: String(r[3] || ''), product: String(r[4] || ''), type: String(r[5] || ''),
    purpose: String(r[6] || ''), dose: String(r[7] || ''), wait: num_(r[8]), ready: ymd_(r[9]), by: String(r[10] || ''),
  }));
}

// ---------- Cosechas ----------
// "Agujeros": los agujeros de cada tanque el día que se anotó la cosecha ("T3=4000; T7=2500"). Quedan fijos:
// si después el dueño cambia los agujeros de un tanque, solo cuenta para las cosechas nuevas.
function holesOf_(names, have) { // names: "Tanque 3, Tanque 7"; have: lo que ya tiene la fila (no se cambia)
  const cfg = (getJSON_('CONFIG', {}) || {}).holes || {};
  const out = {};
  String(have || '').split(/;\s*/).forEach((x) => { const p = x.split('='); if (p.length === 2 && Number(p[1]) > 0) out[p[0].trim()] = Number(p[1]); });
  String(names || '').split(',').map((n) => n.trim()).forEach((n) => {
    const id = Object.keys(TANK_NAMES).find((k) => TANK_NAMES[k] === n);
    if (id && !out[id] && Number(cfg[id]) > 0) out[id] = Math.round(Number(cfg[id]));
  });
  return Object.keys(out).sort().map((k) => k + '=' + out[k]).join('; ');
}
// Al cargar o cambiar agujeros: completa solo los tanques que a una cosecha le faltan (lo que ya tenía no se toca)
function fillHarvHoles_() {
  const sh = ss_().getSheetByName('Cosechas');
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(1, 1, 1, HARV_HEADERS.length).setValues([HARV_HEADERS]);
  const n = sh.getLastRow() - 1, vals = sh.getRange(2, 1, n, HARV_HEADERS.length).getValues();
  let changed = false;
  const col = vals.map((r) => { const h = holesOf_(r[2], r[10]); if (h !== String(r[10] || '')) changed = true; return [h]; });
  if (changed) sh.getRange(2, HARV_HEADERS.length, n, 1).setValues(col);
}
// Una fila por cosecha: el día, el grupo de tanques que se cosechó junto y cuántos paquetes rindió.
// "mine" = la cargó este celular (la puede borrar). El número de dispositivo no sale de la planilla.
// El rendimiento es solo del dueño: a un empleado le llega únicamente lo que anotó él en las últimas 6 horas
// (para que vea que se guardó y pueda corregir un error); después ya no lo ve más.
const HARV_EMP_MS = 6 * 3600000;
function harvFresh_(at) { return at instanceof Date && Date.now() - at.getTime() < HARV_EMP_MS; }
function harvRows_(dev, all) {
  return rows_('Cosechas', HARV_HEADERS.length).slice(-3000).map((r) => ({
    id: String(r[0]), day: ymd_(r[1]), tanks: String(r[2] || ''), packs: num_(r[3]), note: String(r[4] || ''), by: String(r[5] || ''),
    mine: !!dev && String(r[6]) === String(dev), at: iso_(r[7]), fresh: harvFresh_(r[7]),
    boxes: r[8] === '' || r[8] == null ? null : num_(r[8]), detail: String(r[9] || ''), // las cosechas viejas no tienen cajones
    holes: String(r[10] || ''),
  })).filter((x) => x.id && x.day && (all || (x.mine && x.fresh)));
}

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// "2026-11-10" -> fecha al mediodía (así no cambia de día por la zona horaria)
function dayDate_(s) {
  const p = String(s).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12);
}

function iso_(v) {
  return v instanceof Date ? v.toISOString() : String(v || '');
}

// Lee la hoja sin crearla (para las lecturas, que van sin bloqueo)
function rows_(name, cols) {
  const s = ss_().getSheetByName(name);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, cols).getValues();
}

// ---------- Personas y accesos ----------
// PEOPLE: [{ id, name, code, tanks, hours, rate, freq }]. El dueño las crea con un código de 6 números.
// DEVLINK: { dispositivo: id de la persona, o 'OWNER' }. Un celular sin vínculo no puede ver ni guardar nada.
// El código sirve una sola vez: al usarlo queda vinculado ese celular y el código se borra.
function people_() {
  const p = getJSON_('PEOPLE', null);
  if (p) return p;
  // Primera vez: los empleados que ya estaban cargados pasan a ser personas con acceso a Horas
  return getJSON_('EMPLOYEES', []).map((e) => ({ id: Utilities.getUuid(), name: e.name, code: '', tanks: false, hours: true, rate: num_(e.rate), freq: e.freq || '' }));
}

function links_() {
  return getJSON_('DEVLINK', {});
}

// Quién es este dispositivo y qué puede hacer (null = no está vinculado)
function who_(devId) {
  const l = devId ? links_()[String(devId)] : null;
  if (!l) return null;
  if (l === 'OWNER') return { name: 'Dueño', owner: true, tanks: true, hours: true, fum: true, wx: true, harv: true, sow: true, fuel: true };
  const p = people_().find((x) => x.id === l);
  if (!p) return null;
  const fum = p.fum === undefined ? !!p.tanks : !!p.fum;
  // Cosechas: si nunca se tocó, la ve quien entra a Tanques o a Fumigaciones
  // Plantines (siembra): si nunca se tocó, la ve quien entra a Tanques
  return { name: p.name, owner: false, tanks: !!p.tanks, hours: !!p.hours, fum: fum, wx: p.wx !== false, harv: p.harv === true,
    sow: p.sow === undefined ? !!p.tanks : !!p.sow, fuel: p.fuel === true };
}

// ---------- Siembra de plantines (para quien tiene permiso de Plantines) ----------
// Espumas fenólicas: productos cuyo nombre dice "espuma"/"fenólica".
// Semillas: las que el dueño marcó en Proveedores ("Es semilla", va en CONFIG.sowSeed); si nunca se marcó,
// las que se compran por peso y el nombre dice semilla, lechuga, rúcula, etc. (la misma regla está en la app).
// Cada siembra son filas "Ajuste" en la hoja Stock con el detalle "Siembra de plantines: N bandejas · ... · cargó Nombre".
const SOW_FOAM_RE = /espuma|fen[oó]l/i;
const SOW_SEED_RE = /semilla|lechuga|r[uú]cula|albahaca|espinaca|acelga|kale|perejil|cilantro|ciboulette|cebolla|mostaza|mizuna|pak ?choi|radicchio|endivia|achicoria|berro|remolacha|brote/i;
const SOW_DETAIL_RE = /^Siembra de plantines: (\d+) bandeja/;
function sowInfo_(me) {
  const items = {}, mark = (getJSON_('CONFIG', null) || {}).sowSeed || {};
  rows_('Productos de proveedores', SPROD_HEADERS.length).forEach((r) => {
    const n = String(r[3] || '').trim(), u = String(r[5] || ''), k = n.toLowerCase();
    if (!n || items[k]) return;
    const unit = u === 'kg' || u === 'g' ? 'kg' : u === 'L' || u === 'ml' ? 'L' : 'unidad', foam = SOW_FOAM_RE.test(n);
    const seed = !foam && (Object.prototype.hasOwnProperty.call(mark, k) ? !!mark[k] : unit === 'kg' && SOW_SEED_RE.test(n));
    if (foam || seed) items[k] = { name: n, unit: unit, qty: 0, known: false, seed: seed };
  });
  const rows = rows_('Stock', STOCK_HEADERS.length), groups = {}, order = [];
  rows.forEach((r) => {
    const it = items[String(r[2] || '').trim().toLowerCase()];
    if (it) { it.known = true; it.qty = String(r[5]) === 'Conteo' ? num_(r[3]) : it.qty + num_(r[3]); }
    const d = String(r[6] || ''), m = d.match(SOW_DETAIL_RE);
    if (!m) return;
    const g = String(r[0]).split('~')[0];
    if (!groups[g]) { groups[g] = { id: g, day: ymd_(r[1]), trays: Number(m[1]), seed: '', seedG: 0, foam: '', placas: 0, detail: d, at: iso_(r[10]) }; order.push(g); }
    if (String(r[4]) === 'kg') { groups[g].seed = String(r[2]); groups[g].seedG += -num_(r[3]) * 1000; } else { groups[g].foam = String(r[2]); groups[g].placas += -num_(r[3]); }
  });
  const log = order.slice(-15).reverse().map((g) => {
    const x = groups[g], by = (x.detail.match(/ · cargó (.+)$/) || [])[1] || '';
    return Object.assign(x, { by: by, mine: me.owner || (by === me.name && Date.now() - new Date(x.at).getTime() < 48 * 3600000) });
  });
  return { ok: true, items: Object.keys(items).map((k) => items[k]), log: log };
}

// Empleados para las horas: las personas con acceso a Horas
function emps_() {
  return people_().filter((p) => p.hours).map((p) => ({ name: p.name, rate: num_(p.rate), freq: p.freq || '' }));
}

function setUserName_(devId, name) {
  const u = users_();
  const r = u.rowOf(String(devId));
  if (r) u.sh.getRange(r, 2).setValue(name);
}

// Horas sin pagar. mine = las cargó este dispositivo (las puede borrar).
function pendingHours_(vals, devId) {
  return vals.filter((r) => r[6] === 'Pendiente').map((r) => ({
    id: String(r[0]), at: iso_(r[1]), day: ymd_(r[2]), emp: String(r[3]), hours: num_(r[4]), job: String(r[5]),
    by: String(r[10]), mine: !!devId && String(r[9]) === String(devId),
  }));
}

function payments_() {
  return rows_('Pagos', PAY_HEADERS.length).filter((r) => r[0] && !/oculto/i.test(r[10])).map((r) => ({
    id: String(r[0]), date: iso_(r[1]), emp: String(r[2]), hours: num_(r[3]), rate: num_(r[4]), total: num_(r[5]),
    from: ymd_(r[6]), to: ymd_(r[7]), count: num_(r[8]), by: String(r[9]), status: String(r[10]),
  }));
}

// Resumen por mes y por persona: horas trabajadas (por día trabajado) y plata pagada (por fecha de pago)
function months_(vals, pays) {
  const m = {};
  const cell = (mo, emp) => {
    m[mo] = m[mo] || {};
    return (m[mo][emp] = m[mo][emp] || { hours: 0, pending: 0, paid: 0, paidHours: 0, pays: 0 });
  };
  vals.forEach((r) => {
    const d = ymd_(r[2]);
    if (!/^\d{4}-\d{2}/.test(d)) return;
    const c = cell(d.slice(0, 7), String(r[3]));
    c.hours += num_(r[4]);
    if (r[6] === 'Pendiente') c.pending += num_(r[4]);
  });
  pays.forEach((p) => {
    if (/^Anulado/.test(p.status) || !p.date) return;
    const c = cell(ymd_(new Date(p.date)).slice(0, 7), p.emp);
    c.paid += p.total; c.paidHours += p.hours; c.pays++;
  });
  return m;
}

// Todo lo que ve el dueño en Horas del personal (se devuelve también después de cada cambio, así no hace falta otro viaje)
function hoursAdmin_(devId) {
  const pays = payments_();
  const vals = rows_('Horas', HOURS_HEADERS.length);
  return {
    ok: true,
    emps: emps_(),
    pending: pendingHours_(vals, devId),
    payments: pays.slice(-40).reverse(),
    last: lastPaid_(pays),
    months: months_(vals, pays),
  };
}

// Todas las cargas de una persona (pagadas o no), para que el dueño las corrija
function empEntries_(emp) {
  return rows_('Horas', HOURS_HEADERS.length).filter((r) => String(r[3]) === String(emp)).map((r) => ({
    id: String(r[0]), at: iso_(r[1]), day: ymd_(r[2]), emp: String(r[3]), hours: num_(r[4]), job: String(r[5]),
    status: String(r[6]), paidAt: iso_(r[7]), pay: String(r[8]), by: String(r[10]),
  }));
}

// Respuesta para el dueño después de un cambio: todo actualizado y, si se pidió, las cargas de esa persona
function adminReply_(b) {
  const r = hoursAdmin_(b.id);
  if (b.emp) r.entries = empEntries_(b.emp);
  return r;
}

// Recalcula un pago a partir de sus horas (después de corregir o borrar horas ya pagadas).
// Si ya no le queda ninguna hora, el pago se borra.
function fixPay_(pid) {
  if (!pid) return;
  const ps = ss_().getSheetByName('Pagos');
  if (!ps) return;
  const pays = rows_('Pagos', PAY_HEADERS.length);
  const k = pays.findIndex((r) => String(r[0]) === String(pid));
  if (k < 0) return;
  const hrs = rows_('Horas', HOURS_HEADERS.length).filter((r) => String(r[8]) === String(pid) && r[6] === 'Pagado');
  if (!hrs.length) { ps.deleteRow(k + 2); return; }
  let h = 0, from = '', to = '';
  hrs.forEach((r) => { h += num_(r[4]); const d = ymd_(r[2]); if (!from || d < from) from = d; if (!to || d > to) to = d; });
  const rate = num_(pays[k][4]);
  ps.getRange(k + 2, 4, 1, 6).setValues([[h, rate || '', rate ? Math.round(h * rate) : '', dayDate_(from), dayDate_(to), hrs.length]]);
}

// Borra de una hoja las filas que cumplen test (reescribe la hoja de una vez: es mucho más rápido que borrar fila por fila)
function dropRows_(name, cols, test) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return 0;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
  const keep = vals.filter((r) => !test(r));
  if (keep.length === vals.length) return 0;
  sh.getRange(2, 1, vals.length, cols).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, cols).setValues(keep);
  return vals.length - keep.length;
}

// Personas con sus celulares vinculados
function peopleAdmin_() {
  const u = users_();
  const links = links_();
  const devs = (pid) => Object.keys(links).filter((d) => links[d] === pid).map((d) => {
    const r = u.rowOf(d);
    return { id: d, device: r ? String(u.data[r - 1][2]) : '', last: r ? iso_(u.data[r - 1][4]) : '' };
  });
  return { ok: true, people: people_().map((p) => Object.assign({}, p, { devices: devs(p.id) })), owners: devs('OWNER') };
}

// Último pago de cada empleado (sin contar los anulados)
function lastPaid_(pays) {
  const last = {};
  pays.forEach((p) => {
    if (/^Anulado/.test(p.status)) return;
    if (!last[p.emp] || p.date > last[p.emp].date) last[p.emp] = { date: p.date, day: ymd_(new Date(p.date)), hours: p.hours };
  });
  return last;
}

function doGet() {
  return out_({ ok: true, app: 'hidro-dosis' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.action === 'ping') return out_({ ok: true });

  // Un dispositivo bloqueado no lee ni guarda nada (en las acciones de ajustes, b.id es el dispositivo a modificar)
  if (ADMIN_ACTIONS.indexOf(b.action) < 0 && isBlocked_(b.id)) return out_({ ok: false, error: 'blocked' });

  const me = ADMIN_ACTIONS.indexOf(b.action) < 0 ? who_(b.id) : null;
  const noAccess = () => out_({ ok: false, error: 'noaccess', who: me });

  // Lectura rápida, sin bloqueo. Sin vínculo o sin permiso de tanques, no se mandan los tanques.
  if (b.action === 'state') {
    return out_({ ok: true, who: me, config: getJSON_('CONFIG', null), last: me && me.tanks ? getJSON_('LAST', {}) : {}, me: me ? me.name : '' });
  }

  if (b.action === 'history') {
    if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });
    const s = ss_().getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const n = s.getLastRow();
    // Con "tank": todo lo de ese tanque (hasta 3000 movimientos), para ver producciones enteras
    const one = b.tank ? str_(b.tank, 30) : '';
    const start = Math.max(2, n - (one ? 19999 : 399));
    let vals = s.getRange(start, 1, n - start + 1, COLS).getValues();
    if (one) vals = vals.filter((r) => r[1] === one).slice(-3000);
    return out_({
      ok: true,
      tank: one,
      rows: vals.reverse().map((r) => ({ time: r[0], tank: r[1], user: r[2], liters: r[3], ppm: r[4], target: r[5], stage: r[6], water: r[7], summary: r[8], status: r[9], ph: r[10], note: r[11] })),
    });
  }

  // Registro de un día (para todos): movimientos entre from y to
  // Fumigaciones: las ve quien tiene permiso de Fumigaciones (para saber cuándo se puede cosechar)
  if (b.action === 'fum') {
    if (!me || !me.fum) return noAccess();
    return out_({ ok: true, fum: fumRows_() });
  }

  // Cosechas: el dueño ve todas (y el rendimiento); quien tiene permiso de Cosechas, solo las suyas
  if (b.action === 'harv') {
    if (!me || !me.harv) return noAccess();
    return out_({ ok: true, harv: harvRows_(b.id, me.owner) });
  }

  // Siembra de plantines: espumas y semillas con lo que queda, y las últimas siembras
  if (b.action === 'sowinfo') {
    if (!me || !me.sow) return noAccess();
    return out_(sowInfo_(me));
  }

  if (b.action === 'day') {
    if (!me || !me.tanks) return noAccess();
    const s = ss_().getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const from = new Date(b.from), to = new Date(b.to);
    const n = s.getLastRow();
    const start = Math.max(2, n - 1499);
    const vals = s.getRange(start, 1, n - start + 1, COLS).getValues();
    return out_({
      ok: true,
      rows: vals.filter((r) => r[0] instanceof Date && r[0] >= from && r[0] < to)
        .map((r) => ({ time: r[0], tank: r[1], user: r[2], liters: r[3], ppm: r[4], target: r[5], water: r[7], summary: r[8], status: r[9], note: r[11] })),
    });
  }

  // Horas sin pagar y lista de empleados (sin el valor de la hora, que lo ve solo el dueño)
  // Cada persona ve solo lo suyo; el dueño ve a todos
  if (b.action === 'hours') {
    if (!me || !me.hours) return noAccess();
    const last = lastPaid_(payments_());
    const mine = (n) => me.owner || n === me.name;
    const myLast = {};
    if (last[me.name]) myLast[me.name] = last[me.name];
    return out_({
      ok: true,
      emps: me.owner ? emps_().map((e) => e.name) : [me.name],
      pending: pendingHours_(rows_('Horas', HOURS_HEADERS.length), b.id).filter((x) => mine(x.emp)),
      last: me.owner ? last : myLast,
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Cargar horas. El ID lo arma el celular: si llega dos veces (reintento), se guarda una sola.
    // Vincular este celular con el código que le dio el dueño
    if (b.action === 'link') {
      const cache = CacheService.getScriptCache();
      const fails = Number(cache.get('linkfails') || 0);
      if (fails >= 10) return out_({ ok: false, error: 'wait' });
      const code = String(b.code || '').replace(/\D/g, '');
      const people = people_();
      const p = code.length === 6 ? people.find((x) => x.code === code) : null;
      if (!p) { cache.put('linkfails', String(fails + 1), 600); return out_({ ok: false, error: 'code' }); }
      const links = links_();
      links[String(b.id)] = p.id;
      p.code = '';
      setJSON_('DEVLINK', links);
      setJSON_('PEOPLE', people);
      setUserName_(b.id, p.name);
      return out_({ ok: true, who: who_(b.id) });
    }

    // El dueño entra con la contraseña de ajustes: este celular queda como del dueño
    if (b.action === 'owner') {
      if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });
      const links = links_();
      links[String(b.id)] = 'OWNER';
      setJSON_('DEVLINK', links);
      setUserName_(b.id, 'Dueño');
      return out_({ ok: true, who: who_(b.id) });
    }

    if (['hoursadd', 'hoursdel'].indexOf(b.action) >= 0 && (!me || !me.hours)) return noAccess();
    if (['measure', 'done', 'undo', 'tick'].indexOf(b.action) >= 0 && (!me || !me.tanks)) return noAccess();

    if (['harvadd', 'harvdel'].indexOf(b.action) >= 0 && (!me || !me.harv)) return noAccess();
    if (['fuelinfo', 'fueladd'].indexOf(b.action) >= 0 && (!me || !me.fuel)) return noAccess();

    // Entregas para el reparto: las cosechas de los últimos 14 días que el dueño ya repartió entre clientes (sin precios)
    const delivs_ = () => {
      const from = new Date(Date.now() - 14 * 86400000), tanks = {}, by = {}, order = [];
      rows_('Cosechas', 3).forEach((r) => { tanks[String(r[0])] = String(r[2] || ''); });
      rows_('Ventas', SALE_HEADERS.length).forEach((r) => {
        if (!(r[2] instanceof Date) || r[2] < from || !(num_(r[5]) > 0)) return;
        const h = String(r[1]);
        if (!by[h]) { by[h] = { day: ymd_(r[2]), tanks: tanks[h] || '', items: [] }; order.push(h); }
        by[h].items.push({ name: String(r[4] || ''), boxes: num_(r[5]) });
      });
      return order.map((h) => by[h]).sort((a, b) => (a.day < b.day ? 1 : -1));
    };
    // Combustible (chofer o dueño): los km del tablero de la última carga y las últimas cargas
    const fuelInfo = () => {
      const rows = rows_('Vehículo', VEH_HEADERS.length).map((r) => ({ id: String(r[0]), day: ymd_(r[1]), type: String(r[2] || ''), amount: num_(r[3]), liters: num_(r[4]), km: num_(r[5]), detail: String(r[6] || '') }));
      const km = rows.filter((x) => x.km > 0).sort((a, b) => (a.day < b.day ? -1 : a.day > b.day ? 1 : a.km - b.km)).pop();
      return { ok: true, lastKm: km ? km.km : 0, loads: rows.filter((x) => x.type === 'Combustible').sort((a, b) => (a.day < b.day ? 1 : -1)).slice(0, 8)
        .map((x) => ({ day: x.day, amount: x.amount, liters: x.liters, km: x.km, by: (x.detail.match(/^cargó (.+)$/) || [])[1] || '' })), deliv: delivs_() };
    };
    if (b.action === 'fuelinfo') return out_(fuelInfo());
    if (b.action === 'fueladd') {
      const e = b.entry || {}, id = str_(e.id, 40);
      if (!id || !okDay_(e.day) || !(num_(e.amount) > 0) || !(num_(e.km) > 0)) return out_({ ok: false, error: 'fuel' });
      if (!rows_('Vehículo', 1).some((r) => String(r[0]) === id))
        addRows_('Vehículo', VEH_HEADERS, [[id, dayDate_(String(e.day)), 'Combustible', num_(e.amount), num_(e.liters) || '', num_(e.km), 'cargó ' + me.name, new Date()]]);
      return out_(fuelInfo());
    }
    if (['sowadd', 'sowdel'].indexOf(b.action) >= 0 && (!me || !me.sow)) return noAccess();

    // Anotar una siembra de plantines: solo descuenta espumas y semillas (cantidades negativas). El ID lo arma el celular.
    if (b.action === 'sowadd') {
      const valid = {};
      sowInfo_(me).items.forEach((it) => { valid[it.name.toLowerCase()] = it; });
      const have = {};
      rows_('Stock', 1).forEach((r) => { have[String(r[0])] = 1; });
      const now = new Date(), rows = [];
      (Array.isArray(b.moves) ? b.moves : []).slice(0, 4).forEach((m) => {
        const id = str_(m.id, 60), day = /^\d{4}-\d{2}-\d{2}$/.test(String(m.day)) ? String(m.day) : '';
        const it = valid[str_(m.item, 60).trim().toLowerCase()], qty = num_(m.qty), detail = str_(m.detail, 90);
        if (!id || have[id] || !day || !it || !(qty < 0 && qty > -100000) || !SOW_DETAIL_RE.test(detail)) return;
        rows.push([id, dayDate_(day), it.name, qty, it.unit, 'Ajuste', detail + ' · cargó ' + me.name, '', '', '', now]);
        have[id] = 1;
      });
      if (rows.length) { const sh = sheet_('Stock', STOCK_HEADERS); sh.getRange(sh.getLastRow() + 1, 1, rows.length, STOCK_HEADERS.length).setValues(rows); }
      return out_(sowInfo_(me));
    }

    // Borrar una siembra: el dueño cualquiera; los demás, solo las suyas de las últimas 48 horas
    if (b.action === 'sowdel') {
      const g = str_(b.group, 60), row = sowInfo_(me).log.find((x) => x.id === g);
      if (!g || !row) return out_(sowInfo_(me));
      if (!row.mine) return out_({ ok: false, error: 'mine' });
      dropRows_('Stock', STOCK_HEADERS.length, (r) => String(r[0]) === g || String(r[0]).indexOf(g + '~') === 0);
      return out_(sowInfo_(me));
    }

    // Anotar una cosecha. El ID lo arma el celular: si llega dos veces (reintento), se guarda una sola.
    if (b.action === 'harvadd') {
      const e = b.entry || {};
      const id = str_(e.id, 60), day = /^\d{4}-\d{2}-\d{2}$/.test(String(e.day)) ? String(e.day) : '';
      const tanks = (Array.isArray(e.tanks) ? e.tanks : []).map((t) => String(t)).filter((t) => /^T\d$/.test(t) && TANK_NAMES[t]).map((t) => TANK_NAMES[t]);
      // Se anota por cajones: los comunes son de 15 paquetes; si se usaron cajones más grandes, cuántos y de cuántos paquetes.
      // Los paquetes se calculan acá. (Una app vieja todavía puede mandar solo "packs".)
      const int = (v) => { const n = Math.round(num_(v)); return n > 0 ? n : 0; };
      let packs, boxes = '', detail = '';
      if (e.lost) { packs = 0; boxes = 0; detail = 'perdida'; }
      else if (e.boxes !== undefined) {
        const n = int(e.boxes), m = int(e.big), k = m ? int(e.bigSize) : 0;
        if (n > 2000 || m > 2000 || (m && !(k > 0 && k <= 500))) return out_({ ok: false, error: 'harv' });
        packs = n * HARV_BOX + m * k; boxes = n + m;
        if (m) detail = (n ? n + ' de ' + HARV_BOX + ' + ' : '') + m + ' de ' + k;
      } else packs = Math.round(num_(e.packs));
      if (!id || !day || !tanks.length || !(e.lost || (packs > 0 && packs <= 100000))) return out_({ ok: false, error: 'harv' });
      const sh = sheet_('Cosechas', HARV_HEADERS);
      if (!rows_('Cosechas', 1).some((r) => String(r[0]) === id)) {
        sh.getRange(1, 1, 1, HARV_HEADERS.length).setValues([HARV_HEADERS]); // por si la hoja es de antes de los cajones
        sh.appendRow([id, dayDate_(day), tanks.join(', '), packs, str_(e.note, 120).trim(), me.name, str_(b.id, 40), new Date(), boxes, detail, holesOf_(tanks.join(', '), '')]);
        sh.getRange(sh.getLastRow(), 2).setNumberFormat('dd/mm/yyyy');
      }
      return out_({ ok: true, harv: harvRows_(b.id, me.owner) });
    }

    // Borrar una cosecha: el dueño cualquiera; los demás, solo las que cargaron desde su celular en las últimas 6 horas
    if (b.action === 'harvdel') {
      const id = String(b.entry || '');
      const row = rows_('Cosechas', HARV_HEADERS.length).find((r) => String(r[0]) === id);
      if (row && !me.owner && (String(row[6]) !== String(b.id) || !harvFresh_(row[7]))) return out_({ ok: false, error: 'mine' });
      if (row) dropRows_('Cosechas', HARV_HEADERS.length, (r) => String(r[0]) === id);
      return out_({ ok: true, harv: harvRows_(b.id, me.owner) });
    }

    if (b.action === 'hoursadd') {
      const e = b.entry || {};
      const id = str_(e.id, 40), emp = me.owner ? str_(e.emp, 40).trim() : me.name, job = str_(e.job, 60).trim(), hours = num_(e.hours);
      const day = /^\d{4}-\d{2}-\d{2}$/.test(String(e.day)) ? String(e.day) : '';
      if (!id || !emp || !day || !(hours > 0 && hours <= 24)) return out_({ ok: false, error: 'hours' });
      const sh = sheet_('Horas', HOURS_HEADERS);
      if (rows_('Horas', 1).some((r) => String(r[0]) === id)) return out_({ ok: true });
      sh.appendRow([id, b.time ? new Date(b.time) : new Date(), dayDate_(day), emp, hours, job, 'Pendiente', '', '', str_(b.id, 40), users_().nameOf(b.id)]);
      sh.getRange(sh.getLastRow(), 3).setNumberFormat('dd/mm/yyyy');
      return out_({ ok: true });
    }

    // Borrar una carga sin pagar: el celular que la cargó, o el dueño con la contraseña
    if (b.action === 'hoursdel') {
      const vals = rows_('Horas', HOURS_HEADERS.length);
      const i = vals.findIndex((r) => String(r[0]) === String(b.entry));
      const owner = !!b.pin && checkPin_(b.pin);
      if (i < 0) return out_(owner ? adminReply_(b) : { ok: true });
      if (!owner && vals[i][6] !== 'Pendiente') return out_({ ok: false, error: 'paid' });
      if (!owner && String(vals[i][9]) !== String(b.id)) return out_({ ok: false, error: 'pin' });
      ss_().getSheetByName('Horas').deleteRow(i + 2);
      if (vals[i][6] === 'Pagado') fixPay_(vals[i][8]);
      return out_(owner ? adminReply_(b) : { ok: true });
    }

    if (b.action === 'hit') {
      const u = users_();
      const id = str_(b.id, 40);
      if (!id) return out_({ ok: false, error: 'id' });
      const dev = str_(b.device, 60);
      const now = new Date();
      const r = u.rowOf(id);
      let name;
      if (r) {
        name = u.data[r - 1][1];
        u.sh.getRange(r, 3, 1, 4).setValues([[dev, u.data[r - 1][3], now, (Number(u.data[r - 1][5]) || 0) + 1]]);
      } else {
        name = 'Dispositivo ' + u.data.length;
        u.sh.appendRow([id, name, dev, now, now, 1]);
      }
      sheet_('Ingresos', ['Fecha', 'ID', 'Nombre', 'Dispositivo']).appendRow([now, id, name, dev]);
      return out_({ ok: true, me: name });
    }

    const log = () => logSheet_();

    // Nueva medición de un tanque, o tanque apagado (kind = 'off')
    if (b.action === 'measure') {
      const tank = str_(b.tank, 10);
      if (!tank) return out_({ ok: false, error: 'tank' });
      const now = b.time ? new Date(b.time) : new Date();
      const entry = {
        time: now.toISOString(),
        user: users_().nameOf(b.id),
        liters: num_(b.liters),
        ppm: num_(b.ppm),
        target: num_(b.target),
        stage: str_(b.stage, 20),
        fill: !!b.fill,
        kind: str_(b.kind, 10),
        summary: str_(b.summary, 300),
        ph: b.ph ? num_(b.ph) : '',
        note: str_(b.note, 120),
        done: null,
      };
      const last = getJSON_('LAST', {});
      last[tank] = entry;
      setJSON_('LAST', last);
      const estado = entry.kind === 'ok' ? 'En rango' : entry.kind === 'off' ? 'Apagado' : 'Para corregir';
      log().appendRow([now, str_(b.name, 30) || tank, entry.user,
        entry.kind === 'off' ? '' : entry.liters, entry.kind === 'off' ? '' : entry.ppm, entry.kind === 'off' ? '' : entry.target,
        entry.stage, str_(b.water, 30), entry.summary, estado, entry.ph, entry.note]);
      return out_({ ok: true, last: entry });
    }

    // Marcar como corregido
    if (b.action === 'done') {
      const tank = str_(b.tank, 10);
      const last = getJSON_('LAST', {});
      if (!last[tank]) return out_({ ok: false, error: 'tank' });
      const now = b.time ? new Date(b.time) : new Date();
      last[tank].done = { time: now.toISOString(), user: users_().nameOf(b.id) };
      setJSON_('LAST', last);
      log().appendRow([now, str_(b.name, 30) || tank, last[tank].done.user, '', '', '', '', '', last[tank].summary, 'Corregido']);
      return out_({ ok: true, last: last[tank] });
    }

    // Tildar lo que ya se agregó a un tanque para corregir (agua, A, B, C...). Va con la medición: si el tanque
    // se volvió a medir (t distinto), el tilde viejo no se guarda. Lo ven todos los celulares.
    if (b.action === 'tick') {
      const tank = str_(b.tank, 10), k = str_(b.k, 20);
      const last = getJSON_('LAST', {}), e = last[tank];
      if (!e || !k) return out_({ ok: false, error: 'tank' });
      if (String(b.t) !== String(e.time)) return out_({ ok: true, last: e });
      const ticks = e.ticks && typeof e.ticks === 'object' ? e.ticks : {};
      if (b.on) ticks[k] = { time: (b.time ? new Date(b.time) : new Date()).toISOString(), user: users_().nameOf(b.id) };
      else delete ticks[k];
      e.ticks = ticks;
      setJSON_('LAST', last);
      return out_({ ok: true, last: e });
    }

    // Deshacer la última corrección (of = 'done') o la última medición (of = 'measure', con la anterior en prev)
    if (b.action === 'undo') {
      const tank = str_(b.tank, 10);
      if (!tank) return out_({ ok: false, error: 'tank' });
      const last = getJSON_('LAST', {});
      const now = b.time ? new Date(b.time) : new Date();
      const isDone = b.of === 'done';
      if (isDone) { if (last[tank]) last[tank].done = null; }
      else if (b.prev && typeof b.prev === 'object') last[tank] = entry_(b.prev);
      else delete last[tank];
      setJSON_('LAST', last);
      log().appendRow([now, str_(b.name, 30) || tank, users_().nameOf(b.id), '', '', '', '', '',
        isDone ? 'Se deshizo la corrección' : 'Se deshizo la medición', 'Anulado', '', '']);
      return out_({ ok: true, last: last[tank] || null });
    }

    // Desde acá, solo con la contraseña de ajustes
    if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });

    if (b.action === 'sup') return out_(supData_());

    if (b.action === 'fin') return out_(finData_());
    if (b.action === 'finset') {
      const now = new Date();
      // Clientes: alta o precio nuevo (una fila más), o borrar al cliente (todas sus filas)
      const cliDel = (Array.isArray(b.cliDel) ? b.cliDel : []).map(String);
      if (cliDel.length) dropRows_('Clientes', CLI_HEADERS.length, (r) => cliDel.indexOf(String(r[0])) >= 0);
      addRows_('Clientes', CLI_HEADERS, (Array.isArray(b.cliAdd) ? b.cliAdd : []).slice(0, 50)
        .filter((x) => x && str_(x.id, 40) && str_(x.name, 60).trim() && num_(x.price) >= 0 && okDay_(x.since))
        .map((x) => [str_(x.id, 40), str_(x.name, 60).trim(), num_(x.price), dayDate_(String(x.since)), now]));
      // Reparto de una cosecha: se borra el anterior de esa cosecha y se guarda el nuevo
      if (b.sale && str_(b.sale.harvest, 60)) {
        const h = str_(b.sale.harvest, 60);
        dropRows_('Ventas', SALE_HEADERS.length, (r) => String(r[1]) === h);
        addRows_('Ventas', SALE_HEADERS, (Array.isArray(b.sale.rows) ? b.sale.rows : []).slice(0, 30)
          .filter((x) => x && num_(x.boxes) > 0 && okDay_(b.sale.day))
          .map((x, i) => [h + '~' + (i + 1), h, dayDate_(String(b.sale.day)), str_(x.client, 40), str_(x.name, 60), num_(x.boxes), num_(x.packs), num_(x.price), Math.round(num_(x.packs) * num_(x.price)), now]));
      }
      // Gastos del mes y del vehículo: agregar o borrar
      const costDel = (Array.isArray(b.costDel) ? b.costDel : []).map(String);
      if (costDel.length) dropRows_('Gastos del mes', FCOST_HEADERS.length, (r) => costDel.indexOf(String(r[0])) >= 0);
      addRows_('Gastos del mes', FCOST_HEADERS, (Array.isArray(b.costAdd) ? b.costAdd : []).slice(0, 20)
        .filter((x) => x && str_(x.id, 40) && /^\d{4}-\d{2}$/.test(String(x.month)) && num_(x.amount) > 0)
        .map((x) => [str_(x.id, 40), String(x.month), str_(x.type, 30), str_(x.detail, 80), num_(x.amount), now]));
      const vehDel = (Array.isArray(b.vehDel) ? b.vehDel : []).map(String);
      if (vehDel.length) dropRows_('Vehículo', VEH_HEADERS.length, (r) => vehDel.indexOf(String(r[0])) >= 0);
      addRows_('Vehículo', VEH_HEADERS, (Array.isArray(b.vehAdd) ? b.vehAdd : []).slice(0, 20)
        .filter((x) => x && str_(x.id, 40) && okDay_(x.day) && VEH_TYPES.indexOf(String(x.type)) >= 0 && num_(x.amount) >= 0)
        .map((x) => [str_(x.id, 40), dayDate_(String(x.day)), String(x.type), num_(x.amount), num_(x.liters) || '', num_(x.km) || '', str_(x.detail, 80), now]));
      return out_(finData_());
    }

    if (b.action === 'stock') return out_(stockData_());

    // Agregar movimientos de stock. Una compra también deja el precio en "Precios" (mismo ID), así no se carga dos veces.
    // El ID lo arma el celular: si llega dos veces, se guarda una sola.
    if (b.action === 'stockadd') {
      const have = {};
      rows_('Stock', 1).forEach((r) => { have[String(r[0])] = 1; });
      const prods = {};
      rows_('Productos de proveedores', SPROD_HEADERS.length).forEach((r) => { prods[String(r[0])] = { id: String(r[0]), sup: String(r[2] || ''), name: String(r[3] || '') }; });
      const now = new Date(), rows = [], prices = [];
      (Array.isArray(b.moves) ? b.moves : []).slice(0, 60).forEach((m) => {
        const id = str_(m.id, 60), day = /^\d{4}-\d{2}-\d{2}$/.test(String(m.day)) ? String(m.day) : '';
        const item = str_(m.item, 60).trim(), type = STOCK_TYPES.indexOf(String(m.type)) >= 0 ? String(m.type) : '';
        const unit = STOCK_UNITS.indexOf(String(m.unit)) >= 0 ? String(m.unit) : '';
        if (!id || have[id] || !day || !item || !type || !unit || !isFinite(Number(m.qty))) return;
        const p = m.product ? prods[String(m.product)] : null;
        rows.push([id, dayDate_(day), item, num_(m.qty), unit, type, str_(m.detail, 120), p ? p.sup : '', num_(m.price) || '', num_(m.total) || '', now]);
        if (type === 'Compra' && p && num_(m.price) > 0) prices.push([id, p.id, p.name, p.sup, dayDate_(day), num_(m.price)]);
        have[id] = 1;
      });
      if (rows.length) { const sh = sheet_('Stock', STOCK_HEADERS); sh.getRange(sh.getLastRow() + 1, 1, rows.length, STOCK_HEADERS.length).setValues(rows); }
      if (prices.length) { const sh = sheet_('Precios', PRICE_HEADERS); sh.getRange(sh.getLastRow() + 1, 1, prices.length, PRICE_HEADERS.length).setValues(prices); }
      if (b.orderDone) setOrderState_(str_(b.orderDone, 60), 'Recibido'); // la compra de un pedido que llegó
      return out_(stockData_());
    }

    // Anotar un pedido hecho a un proveedor. El ID lo arma el celular: si llega dos veces, se guarda una sola.
    if (b.action === 'orderadd') {
      const o = b.order || {}, id = str_(o.id, 60), day = /^\d{4}-\d{2}-\d{2}$/.test(String(o.day)) ? String(o.day) : '';
      const sups = {};
      rows_('Proveedores', 2).forEach((r) => { sups[String(r[0])] = String(r[1] || ''); });
      const items = (Array.isArray(o.items) ? o.items : []).slice(0, 60).filter((it) => it && str_(it.item, 60).trim() && num_(it.n) > 0);
      if (!id || id.indexOf('~') >= 0 || !day || !sups[String(o.sup)] || !items.length) return out_({ ok: false, error: 'order' });
      if (!rows_('Pedidos', 1).some((r) => String(r[0]).split('~')[0] === id)) {
        const now = new Date(), sh = sheet_('Pedidos', ORDER_HEADERS);
        sh.getRange(sh.getLastRow() + 1, 1, items.length, ORDER_HEADERS.length).setValues(items.map((it, i) => [
          id + '~' + (i + 1), dayDate_(day), String(o.sup), sups[String(o.sup)], str_(it.item, 60).trim(), str_(it.product, 40), num_(it.n), str_(it.detail, 80), 'Pedido', now,
        ]));
      }
      return out_(stockData_());
    }

    // Marcar un pedido como Recibido o Cancelado
    if (b.action === 'orderset') {
      const st = ORDER_STATES.indexOf(String(b.status)) >= 0 ? String(b.status) : '';
      if (!st) return out_({ ok: false, error: 'order' });
      setOrderState_(str_(b.group, 60), st);
      return out_(stockData_());
    }

    // Borrar movimientos cargados por error (todos los de una misma carga comparten el comienzo del ID)
    if (b.action === 'stockdel') {
      const g = str_(b.group, 60);
      if (g) {
        const test = (r) => String(r[0]) === g || String(r[0]).indexOf(g + '~') === 0;
        dropRows_('Stock', STOCK_HEADERS.length, test);
        dropRows_('Precios', PRICE_HEADERS.length, test);
      }
      return out_(stockData_());
    }

    // Guardar proveedores y productos (enteros), agregar o borrar precios, y la receta de costos
    if (b.action === 'supset') {
      const sups = (Array.isArray(b.suppliers) ? b.suppliers : []).slice(0, 200)
        .map((s) => ({ id: str_(s.id, 40), name: str_(s.name, 60).trim(), phone: str_(s.phone, 40), notes: str_(s.notes, 200) }))
        .filter((s) => s.id && s.name);
      const supName = {};
      sups.forEach((s) => { supName[s.id] = s.name; });
      const prods = (Array.isArray(b.products) ? b.products : []).slice(0, 500)
        .map((p) => ({ id: str_(p.id, 40), sup: str_(p.sup, 40), name: str_(p.name, 60).trim(), brand: str_(p.brand, 40).trim(), size: num_(p.size), unit: PROD_UNITS.indexOf(String(p.unit)) >= 0 ? String(p.unit) : 'unidad' }))
        .filter((p) => p.id && p.name && supName[p.sup]);
      const prodById = {};
      prods.forEach((p) => { prodById[p.id] = p; });
      writeAll_('Proveedores', SUP_HEADERS, sups.map((s) => [s.id, s.name, s.phone, s.notes]));
      writeAll_('Productos de proveedores', SPROD_HEADERS, prods.map((p) => [p.id, p.sup, supName[p.sup], p.name, p.size, p.unit, p.brand]));
      const del = (Array.isArray(b.delPrices) ? b.delPrices : []).map(String);
      if (del.length) dropRows_('Precios', PRICE_HEADERS.length, (r) => del.indexOf(String(r[0])) >= 0);
      const add = (Array.isArray(b.addPrices) ? b.addPrices : []).filter((x) => x && prodById[String(x.product)] && /^\d{4}-\d{2}-\d{2}$/.test(String(x.day)) && num_(x.price) > 0);
      if (add.length) {
        const have = {};
        rows_('Precios', 1).forEach((r) => { have[String(r[0])] = 1; });
        const rows = add.filter((x) => !have[str_(x.id, 40)]).map((x) => {
          const p = prodById[String(x.product)];
          return [str_(x.id, 40), p.id, p.name, supName[p.sup], dayDate_(String(x.day)), num_(x.price)];
        });
        if (rows.length) {
          const sh = sheet_('Precios', PRICE_HEADERS);
          sh.getRange(sh.getLastRow() + 1, 1, rows.length, PRICE_HEADERS.length).setValues(rows);
        }
      }
      if (b.recipe && typeof b.recipe === 'object') setJSON_('COSTRECIPE', b.recipe);
      return out_(supData_());
    }

    // Cargar una fumigación con uno o más productos. El ID lo arma el celular: si llega dos veces, se guarda una sola.
    if (b.action === 'fumadd') {
      const e = b.entry || {};
      const group = str_(e.id, 40), day = /^\d{4}-\d{2}-\d{2}$/.test(String(e.day)) ? String(e.day) : '';
      const tanks = (Array.isArray(e.tanks) ? e.tanks : []).map((t) => TANK_NAMES[String(t)]).filter(Boolean);
      const prods = (Array.isArray(e.products) ? e.products : []).slice(0, 10).map((p) => ({
        product: str_(p.product, 60).trim(), type: str_(p.type, 20), purpose: str_(p.purpose, 80), dose: str_(p.dose, 60), wait: Math.round(num_(p.wait)),
      }));
      if (!group || !day || !tanks.length || !prods.length || prods.some((p) => !p.product || p.wait < 0 || p.wait > 365)) return out_({ ok: false, error: 'fum' });
      const sh = sheet_('Fumigaciones', FUM_HEADERS);
      if (!rows_('Fumigaciones', 2).some((r) => String(r[1]) === group)) {
        const now = new Date();
        sh.getRange(sh.getLastRow() + 1, 1, prods.length, FUM_HEADERS.length).setValues(prods.map((p, i) => {
          const ready = dayDate_(day);
          ready.setDate(ready.getDate() + p.wait);
          return [group + '-' + (i + 1), group, dayDate_(day), tanks.join(', '), p.product, p.type, p.purpose, p.dose, p.wait, ready, 'Dueño', now];
        }));
      }
      return out_({ ok: true, fum: fumRows_() });
    }

    // Borrar una fumigación cargada por error (todos sus productos)
    if (b.action === 'fumdel') {
      dropRows_('Fumigaciones', FUM_HEADERS.length, (r) => String(r[1]) === String(b.entry));
      return out_({ ok: true, fum: fumRows_() });
    }

    if (b.action === 'setconfig') {
      if (!b.config || typeof b.config !== 'object') return out_({ ok: false, error: 'config' });
      setJSON_('CONFIG', b.config);
      if (b.config.holes) fillHarvHoles_();
      return out_({ ok: true });
    }

    if (b.action === 'list') {
      const u = users_();
      const blocked = blockedIds_();
      const users = u.data.slice(1).map((r) => ({ id: r[0], name: r[1], device: r[2], first: r[3], last: r[4], count: r[5], blocked: blocked.indexOf(String(r[0])) >= 0 }));
      // Bloqueados que ya se borraron de la lista: se muestran para poder desbloquearlos
      blocked.forEach((id) => {
        if (!u.rowOf(id)) users.push({ id: id, name: '', device: '', first: '', last: '', count: '', blocked: true, gone: true });
      });
      return out_({ ok: true, users: users });
    }

    // Bloquear o desbloquear un dispositivo
    if (b.action === 'block') {
      const id = str_(b.id, 40);
      if (!id) return out_({ ok: false, error: 'id' });
      const ids = blockedIds_().filter((x) => x !== id);
      if (b.blocked) ids.push(id);
      setJSON_('BLOCKED', ids);
      return out_({ ok: true });
    }

    // Borrar un dispositivo de la lista (si estaba bloqueado, sigue bloqueado)
    if (b.action === 'remove') {
      const u = users_();
      const r = u.rowOf(String(b.id));
      if (r) u.sh.deleteRow(r);
      return out_({ ok: true });
    }

    // Horas para el dueño: todo lo pendiente, empleados con su valor de hora y los últimos pagos
    if (b.action === 'hoursadmin') return out_(hoursAdmin_(b.id));

    // Personas: lista con sus celulares vinculados
    if (b.action === 'people') return out_(peopleAdmin_());

    // Guardar personas. Si a alguien le cambian el nombre, sus horas sin pagar pasan al nombre nuevo.
    if (b.action === 'setpeople') {
      if (!Array.isArray(b.people)) return out_({ ok: false, error: 'people' });
      const old = people_();
      const seen = {}, codes = {};
      const list = [];
      b.people.forEach((x) => {
        const name = str_(x && x.name, 40).trim();
        if (!name || seen[name.toLowerCase()]) return;
        let code = String(x.code || '').replace(/\D/g, '').slice(0, 6);
        if (code.length !== 6 || codes[code]) code = '';
        seen[name.toLowerCase()] = true;
        if (code) codes[code] = true;
        list.push({ id: str_(x.id, 40) || Utilities.getUuid(), name: name, code: code, tanks: !!x.tanks, hours: !!x.hours, fum: x.fum === undefined ? !!x.tanks : !!x.fum, wx: x.wx !== false,
          harv: x.harv === true, sow: x.sow === undefined ? !!x.tanks : !!x.sow, fuel: x.fuel === true,
          rate: num_(x.rate), freq: x.freq === 'quincenal' ? 'quincenal' : x.freq === 'semanal' ? 'semanal' : '' });
      });
      const renames = {};
      old.forEach((o) => { const n = list.find((p) => p.id === o.id); if (n && n.name !== o.name) renames[o.name] = n.name; });
      const vals = rows_('Horas', HOURS_HEADERS.length);
      if (Object.keys(renames).length && vals.length) {
        vals.forEach((r) => { if (r[6] === 'Pendiente' && renames[r[3]]) r[3] = renames[r[3]]; });
        ss_().getSheetByName('Horas').getRange(2, 4, vals.length, 1).setValues(vals.map((r) => [r[3]]));
      }
      // Personas quitadas: se borran todas sus horas y sus pagos
      old.forEach((o) => {
        if (list.some((p) => p.id === o.id || p.name === o.name)) return;
        dropRows_('Horas', HOURS_HEADERS.length, (r) => String(r[3]) === o.name);
        dropRows_('Pagos', PAY_HEADERS.length, (r) => String(r[2]) === o.name);
      });
      // Personas quitadas: sus celulares quedan sin vínculo
      const links = links_();
      Object.keys(links).forEach((d) => { if (links[d] !== 'OWNER' && !list.some((p) => p.id === links[d])) delete links[d]; });
      setJSON_('DEVLINK', links);
      setJSON_('PEOPLE', list);
      return out_(peopleAdmin_());
    }

    // Desvincular un celular (tiene que volver a poner un código)
    if (b.action === 'unlink') {
      const links = links_();
      delete links[String(b.dev)];
      setJSON_('DEVLINK', links);
      return out_(peopleAdmin_());
    }

    // Pagar: las cargas que el dueño vio (ids) pasan a "Pagado" y se anota el pago
    if (b.action === 'hourspay') {
      const emp = str_(b.emp, 40);
      const ids = (Array.isArray(b.ids) ? b.ids : []).map(String);
      const sh = sheet_('Horas', HOURS_HEADERS);
      const vals = rows_('Horas', HOURS_HEADERS.length);
      const hit = [];
      vals.forEach((r, i) => { if (r[6] === 'Pendiente' && String(r[3]) === emp && ids.indexOf(String(r[0])) >= 0) hit.push(i); });
      if (!hit.length) return out_({ ok: false, error: 'none' });
      const now = new Date();
      const pid = Utilities.getUuid();
      const e = emps_().find((x) => x.name === emp) || {};
      const rate = num_(e.rate);
      let hours = 0, from = '', to = '';
      hit.forEach((i) => {
        hours += num_(vals[i][4]);
        const d = ymd_(vals[i][2]);
        if (!from || d < from) from = d;
        if (!to || d > to) to = d;
      });
      const cells = (col) => hit.map((i) => col + (i + 2));
      sh.getRangeList(cells('G')).setValue('Pagado');
      sh.getRangeList(cells('H')).setValue(now);
      sh.getRangeList(cells('I')).setValue(pid);
      const total = Math.round(hours * rate);
      const ps = sheet_('Pagos', PAY_HEADERS);
      ps.appendRow([pid, now, emp, hours, rate || '', rate ? total : '', dayDate_(from), dayDate_(to), hit.length, users_().nameOf(b.id), 'Pagado']);
      ps.getRange(ps.getLastRow(), 7, 1, 2).setNumberFormat('dd/mm/yyyy');
      return out_(Object.assign(adminReply_(b), { paid: { hours: hours, total: total } }));
    }

    // Borrar pagos: desaparecen de la lista, del resumen mensual y de la hoja Pagos. Sus horas siguen como pagadas.
    if (b.action === 'paydel') {
      const ids = (Array.isArray(b.pays) ? b.pays : []).map(String);
      dropRows_('Pagos', PAY_HEADERS.length, (r) => ids.indexOf(String(r[0])) >= 0);
      return out_(adminReply_(b));
    }

    // Borrar todas las horas y pagos de alguien que ya no está en Personas (datos que quedaron sueltos)
    if (b.action === 'empwipe') {
      const emp = str_(b.emp, 40);
      if (!emp || people_().some((p) => p.name === emp)) return out_({ ok: false, error: 'person' });
      dropRows_('Horas', HOURS_HEADERS.length, (r) => String(r[3]) === emp);
      dropRows_('Pagos', PAY_HEADERS.length, (r) => String(r[2]) === emp);
      return out_(hoursAdmin_(b.id));
    }

    // Lista de todas las cargas de una persona
    if (b.action === 'empentries') return out_({ ok: true, entries: empEntries_(str_(b.emp, 40)) });

    // Corregir una carga (día, horas, trabajo). Si ya estaba pagada, se recalcula ese pago.
    if (b.action === 'hoursedit') {
      const vals = rows_('Horas', HOURS_HEADERS.length);
      const i = vals.findIndex((r) => String(r[0]) === String(b.entry));
      if (i < 0) return out_({ ok: false, error: 'gone' });
      const day = /^\d{4}-\d{2}-\d{2}$/.test(String(b.day)) ? String(b.day) : '';
      const hours = num_(b.hours), job = str_(b.job, 60).trim();
      if (!day || !(hours > 0 && hours <= 24) || !job) return out_({ ok: false, error: 'hours' });
      const sh = ss_().getSheetByName('Horas');
      sh.getRange(i + 2, 3).setValue(dayDate_(day));
      sh.getRange(i + 2, 5, 1, 2).setValues([[hours, job]]);
      if (vals[i][6] === 'Pagado') fixPay_(vals[i][8]);
      return out_(adminReply_(b));
    }

    // Deshacer un pago: sus horas vuelven a "Pendiente" y el pago queda "Anulado"
    if (b.action === 'hoursunpay') {
      const pid = str_(b.pay, 40);
      const ps = ss_().getSheetByName('Pagos');
      const pays = rows_('Pagos', 1);
      const k = pays.findIndex((r) => String(r[0]) === pid);
      if (!pid || k < 0) return out_({ ok: false, error: 'pay' });
      const hit = [];
      rows_('Horas', HOURS_HEADERS.length).forEach((r, i) => { if (String(r[8]) === pid) hit.push(i + 2); });
      if (hit.length) {
        const sh = ss_().getSheetByName('Horas');
        sh.getRangeList(hit.map((n) => 'G' + n)).setValue('Pendiente');
        sh.getRangeList(hit.map((n) => 'H' + n + ':I' + n)).setValue('');
      }
      ps.getRange(k + 2, PAY_HEADERS.length).setValue('Anulado');
      return out_(adminReply_(b));
    }

    if (b.action === 'rename') {
      const u = users_();
      const r = u.rowOf(String(b.id));
      if (!r) return out_({ ok: false, error: 'id' });
      u.sh.getRange(r, 2).setValue(str_(b.name, 40));
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Avisos con la app cerrada: Telegram (o mail) ----------
// Avisa cuando un tanque lleva más horas sin corregir que las que dicen los Ajustes de la app
// ("Avisar después de"). Con 0 no avisa.
//
// Telegram: guardá dos datos en Configuración del proyecto > Propiedades del script:
//   TELEGRAM_TOKEN  el token que te da @BotFather al crear el bot (es secreto: no lo pegues en ningún otro lado)
//   TELEGRAM_CHAT   el número de chat que tiene que recibir los avisos (varios, separados por coma)
// Para saber el número de chat: escribile "hola" al bot desde Telegram y ejecutá verChatsTelegram.
// Si no hay datos de Telegram, el aviso sale por mail a la cuenta dueña del script.
//
// Activar: ejecutá instalarAvisos una sola vez y aceptá los permisos. Probar: probarAviso. Apagar: quitarAvisos.

function telegram_() {
  const token = props_().getProperty('TELEGRAM_TOKEN');
  const chats = String(props_().getProperty('TELEGRAM_CHAT') || '').split(',').map((c) => c.trim()).filter(Boolean);
  return token && chats.length ? { token: token, chats: chats } : null;
}

// Manda el aviso por Telegram si está configurado; si no, por mail. Devuelve por dónde salió.
function avisar_(titulo, texto) {
  const tg = telegram_();
  if (!tg) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Hidroponia: ' + titulo, texto);
    return 'mail';
  }
  tg.chats.forEach((chat) => {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + tg.token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chat, text: '💧 ' + titulo + '\n\n' + texto }),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) throw new Error('Telegram no aceptó el mensaje para el chat ' + chat + ' (código ' + res.getResponseCode() + '). Revisá TELEGRAM_TOKEN y TELEGRAM_CHAT.');
  });
  return 'Telegram';
}

function avisoPendientes() {
  const cfg = getJSON_('CONFIG', {}) || {};
  const horas = num_(cfg.alertH);
  if (!horas) return;
  const last = getJSON_('LAST', {});
  const avisados = getJSON_('AVISADOS', {});
  const ahora = Date.now();
  const tarde = [];
  Object.keys(last).forEach((id) => {
    const l = last[id];
    if (!l || l.done || l.kind === 'ok' || l.kind === 'off') return;
    if (ahora - new Date(l.time).getTime() < horas * 3600000) return;
    if (avisados[id] === l.time) return; // ya se avisó por esta medición
    avisados[id] = l.time;
    tarde.push('• ' + (TANK_NAMES[id] || id) + ': ' + l.summary + ' (medido por ' + l.user + ')');
  });
  if (!tarde.length) return;
  avisar_(tarde.length === 1 ? '1 tanque sin corregir' : tarde.length + ' tanques sin corregir',
    'Llevan más de ' + horas + ' h sin corregir:\n' + tarde.join('\n'));
  setJSON_('AVISADOS', avisados); // se anota recién cuando el aviso salió bien
}

// Manda un aviso de prueba ahora mismo y dice por dónde salió.
function probarAviso() {
  const via = avisar_('aviso de prueba', 'Si estás leyendo esto, los avisos de la app funcionan.');
  Logger.log('Aviso de prueba enviado por ' + via + '.');
}

// Muestra en el registro de ejecución los chats que le escribieron al bot, con su número.
function verChatsTelegram() {
  const token = props_().getProperty('TELEGRAM_TOKEN');
  if (!token) { Logger.log('Falta guardar TELEGRAM_TOKEN en las Propiedades del script.'); return; }
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) { Logger.log('Telegram no aceptó el token (código ' + res.getResponseCode() + ').'); return; }
  const vistos = {};
  (JSON.parse(res.getContentText()).result || []).forEach((u) => {
    const c = (u.message || u.channel_post || {}).chat;
    if (c) vistos[c.id] = c.title || [c.first_name, c.last_name].filter(Boolean).join(' ') || c.username || '';
  });
  const ids = Object.keys(vistos);
  if (!ids.length) { Logger.log('Nadie le escribió al bot todavía. Mandale "hola" desde Telegram y volvé a ejecutar.'); return; }
  ids.forEach((id) => Logger.log('TELEGRAM_CHAT = ' + id + '   (' + vistos[id] + ')'));
}

function quitarAvisos() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'avisoPendientes') ScriptApp.deleteTrigger(t);
  });
}

function instalarAvisos() {
  quitarAvisos();
  ScriptApp.newTrigger('avisoPendientes').timeBased().everyMinutes(30).create();
}
