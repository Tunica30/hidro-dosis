// Registro de Dosis Hidroponia: usuarios, mediciones de tanques y ajustes compartidos.
// Proyecto de Apps Script publicado como "Aplicación web".

// Planilla "Hidro Dosis Usuarios"
const SS = SpreadsheetApp.openById('1vxmWFBLpxiWxihmsyGTzwQ2NOoUXO5o54gotwPuVGYA');
const LOG_HEADERS = ['Fecha', 'Tanque', 'Usuario', 'Litros', 'ppm', 'Objetivo', 'Etapa', 'Agua', 'Indicación', 'Estado', 'pH', 'Nota'];
const COLS = LOG_HEADERS.length;
const TANK_NAMES = { T1: 'Tanque 1', T2: 'Tanque 2', T3: 'Tanque 3', T4: 'Tanque 4', T5: 'Tanque 5', T6: 'Tanque 6', T7: 'Tanque 7', T8: 'Tanque 8',
  P1: 'Plantinero 1', P2: 'Plantinero 2', P3: 'Plantinero 3', P4: 'Plantinero 4' };

function sheet_(name, headers) {
  let s = SS.getSheetByName(name);
  if (!s) {
    s = SS.insertSheet(name);
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

const ADMIN_ACTIONS = ['history', 'setconfig', 'list', 'rename', 'block', 'remove', 'people', 'setpeople', 'unlink', 'hoursadmin', 'hourspay', 'hoursunpay'];

// ---------- Horas de trabajo ----------
// Hoja "Horas": una fila por carga. Estado "Pendiente" hasta que el dueño la paga; ahí pasa a "Pagado".
// Hoja "Pagos": una fila por pago (con el total de horas y de plata).
// Empleados: las personas con acceso a Horas (ver "Personas y accesos" más abajo).
const HOURS_HEADERS = ['ID', 'Cargado', 'Día', 'Empleado', 'Horas', 'Trabajo', 'Estado', 'Pagado el', 'Pago', 'Dispositivo', 'Cargado por'];
const PAY_HEADERS = ['ID', 'Fecha', 'Empleado', 'Horas', 'Valor hora', 'Total', 'Desde', 'Hasta', 'Cargas', 'Pagado por', 'Estado'];

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SS.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
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
  const s = SS.getSheetByName(name);
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
  if (l === 'OWNER') return { name: 'Dueño', owner: true, tanks: true, hours: true };
  const p = people_().find((x) => x.id === l);
  return p ? { name: p.name, owner: false, tanks: !!p.tanks, hours: !!p.hours } : null;
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
  return rows_('Pagos', PAY_HEADERS.length).map((r) => ({
    id: String(r[0]), date: iso_(r[1]), emp: String(r[2]), hours: num_(r[3]), rate: num_(r[4]), total: num_(r[5]),
    from: ymd_(r[6]), to: ymd_(r[7]), count: num_(r[8]), by: String(r[9]), status: String(r[10]),
  }));
}

// Último pago de cada empleado (sin contar los anulados)
function lastPaid_(pays) {
  const last = {};
  pays.forEach((p) => {
    if (p.status === 'Anulado') return;
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
    const s = SS.getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const n = s.getLastRow();
    const start = Math.max(2, n - 399);
    const vals = s.getRange(start, 1, n - start + 1, COLS).getValues();
    return out_({
      ok: true,
      rows: vals.reverse().map((r) => ({ time: r[0], tank: r[1], user: r[2], liters: r[3], ppm: r[4], target: r[5], stage: r[6], water: r[7], summary: r[8], status: r[9], ph: r[10], note: r[11] })),
    });
  }

  // Registro de un día (para todos): movimientos entre from y to
  if (b.action === 'day') {
    if (!me || !me.tanks) return noAccess();
    const s = SS.getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const from = new Date(b.from), to = new Date(b.to);
    const n = s.getLastRow();
    const start = Math.max(2, n - 1499);
    const vals = s.getRange(start, 1, n - start + 1, COLS).getValues();
    return out_({
      ok: true,
      rows: vals.filter((r) => r[0] instanceof Date && r[0] >= from && r[0] < to)
        .map((r) => ({ time: r[0], tank: r[1], user: r[2], summary: r[8], status: r[9], note: r[11] })),
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
    if (['measure', 'done', 'undo'].indexOf(b.action) >= 0 && (!me || !me.tanks)) return noAccess();

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
      if (i < 0) return out_({ ok: true });
      if (vals[i][6] !== 'Pendiente') return out_({ ok: false, error: 'paid' });
      if (String(vals[i][9]) !== String(b.id) && !checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });
      SS.getSheetByName('Horas').deleteRow(i + 2);
      return out_({ ok: true });
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

    if (b.action === 'setconfig') {
      if (!b.config || typeof b.config !== 'object') return out_({ ok: false, error: 'config' });
      setJSON_('CONFIG', b.config);
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
    if (b.action === 'hoursadmin') {
      const pays = payments_();
      return out_({
        ok: true,
        emps: emps_(),
        pending: pendingHours_(rows_('Horas', HOURS_HEADERS.length), b.id),
        payments: pays.slice(-40).reverse(),
        last: lastPaid_(pays),
      });
    }

    // Personas: lista con sus celulares vinculados
    if (b.action === 'people') {
      const u = users_();
      const links = links_();
      const devs = (pid) => Object.keys(links).filter((d) => links[d] === pid).map((d) => {
        const r = u.rowOf(d);
        return { id: d, device: r ? String(u.data[r - 1][2]) : '', last: r ? iso_(u.data[r - 1][4]) : '' };
      });
      return out_({ ok: true, people: people_().map((p) => Object.assign({}, p, { devices: devs(p.id) })), owners: devs('OWNER') });
    }

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
        list.push({ id: str_(x.id, 40) || Utilities.getUuid(), name: name, code: code, tanks: !!x.tanks, hours: !!x.hours,
          rate: num_(x.rate), freq: x.freq === 'quincenal' ? 'quincenal' : x.freq === 'semanal' ? 'semanal' : '' });
      });
      const renames = {};
      old.forEach((o) => { const n = list.find((p) => p.id === o.id); if (n && n.name !== o.name) renames[o.name] = n.name; });
      const vals = rows_('Horas', HOURS_HEADERS.length);
      if (Object.keys(renames).length && vals.length) {
        vals.forEach((r) => { if (r[6] === 'Pendiente' && renames[r[3]]) r[3] = renames[r[3]]; });
        SS.getSheetByName('Horas').getRange(2, 4, vals.length, 1).setValues(vals.map((r) => [r[3]]));
      }
      // Personas borradas: sus celulares quedan sin vínculo
      const links = links_();
      Object.keys(links).forEach((d) => { if (links[d] !== 'OWNER' && !list.some((p) => p.id === links[d])) delete links[d]; });
      setJSON_('DEVLINK', links);
      setJSON_('PEOPLE', list);
      return out_({ ok: true });
    }

    // Desvincular un celular (tiene que volver a poner un código)
    if (b.action === 'unlink') {
      const links = links_();
      delete links[String(b.dev)];
      setJSON_('DEVLINK', links);
      return out_({ ok: true });
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
        vals[i][6] = 'Pagado'; vals[i][7] = now; vals[i][8] = pid;
      });
      sh.getRange(2, 7, vals.length, 3).setValues(vals.map((r) => [r[6], r[7], r[8]]));
      const total = Math.round(hours * rate);
      const ps = sheet_('Pagos', PAY_HEADERS);
      ps.appendRow([pid, now, emp, hours, rate || '', rate ? total : '', dayDate_(from), dayDate_(to), hit.length, users_().nameOf(b.id), 'Pagado']);
      ps.getRange(ps.getLastRow(), 7, 1, 2).setNumberFormat('dd/mm/yyyy');
      return out_({ ok: true, hours: hours, total: total });
    }

    // Deshacer un pago: sus horas vuelven a "Pendiente" y el pago queda "Anulado"
    if (b.action === 'hoursunpay') {
      const pid = str_(b.pay, 40);
      const ps = SS.getSheetByName('Pagos');
      const pays = rows_('Pagos', 1);
      const k = pays.findIndex((r) => String(r[0]) === pid);
      if (!pid || k < 0) return out_({ ok: false, error: 'pay' });
      const vals = rows_('Horas', HOURS_HEADERS.length);
      if (vals.length) {
        vals.forEach((r) => { if (String(r[8]) === pid) { r[6] = 'Pendiente'; r[7] = ''; r[8] = ''; } });
        SS.getSheetByName('Horas').getRange(2, 7, vals.length, 3).setValues(vals.map((r) => [r[6], r[7], r[8]]));
      }
      ps.getRange(k + 2, PAY_HEADERS.length).setValue('Anulado');
      return out_({ ok: true });
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
