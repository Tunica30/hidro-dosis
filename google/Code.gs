// Registro de Dosis Hidroponia: usuarios, mediciones de tanques y ajustes compartidos.
// Proyecto de Apps Script publicado como "Aplicación web".

// Planilla "Hidro Dosis Usuarios"
const SS = SpreadsheetApp.openById('1vxmWFBLpxiWxihmsyGTzwQ2NOoUXO5o54gotwPuVGYA');
const LOG_HEADERS = ['Fecha', 'Tanque', 'Usuario', 'Litros', 'ppm', 'Objetivo', 'Etapa', 'Agua', 'Indicación', 'Estado'];

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

function doGet() {
  return out_({ ok: true, app: 'hidro-dosis' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.action === 'ping') return out_({ ok: true });

  // Lectura rápida, sin bloqueo
  if (b.action === 'state') {
    return out_({ ok: true, config: getJSON_('CONFIG', null), last: getJSON_('LAST', {}), me: b.id ? users_().nameOf(b.id) : '' });
  }

  if (b.action === 'history') {
    if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });
    const s = SS.getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const n = s.getLastRow();
    const start = Math.max(2, n - 399);
    const vals = s.getRange(start, 1, n - start + 1, 10).getValues();
    return out_({
      ok: true,
      rows: vals.reverse().map((r) => ({ time: r[0], tank: r[1], user: r[2], liters: r[3], ppm: r[4], target: r[5], stage: r[6], water: r[7], summary: r[8], status: r[9] })),
    });
  }

  // Registro de un día (para todos): movimientos entre from y to
  if (b.action === 'day') {
    const s = SS.getSheetByName('Mediciones');
    if (!s || s.getLastRow() < 2) return out_({ ok: true, rows: [] });
    const from = new Date(b.from), to = new Date(b.to);
    const n = s.getLastRow();
    const start = Math.max(2, n - 1499);
    const vals = s.getRange(start, 1, n - start + 1, 10).getValues();
    return out_({
      ok: true,
      rows: vals.filter((r) => r[0] instanceof Date && r[0] >= from && r[0] < to)
        .map((r) => ({ time: r[0], tank: r[1], user: r[2], summary: r[8], status: r[9] })),
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
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

    const log = () => sheet_('Mediciones', LOG_HEADERS);

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
        done: null,
      };
      const last = getJSON_('LAST', {});
      last[tank] = entry;
      setJSON_('LAST', last);
      const estado = entry.kind === 'ok' ? 'En rango' : entry.kind === 'off' ? 'Apagado' : 'Para corregir';
      log().appendRow([now, str_(b.name, 30) || tank, entry.user,
        entry.kind === 'off' ? '' : entry.liters, entry.kind === 'off' ? '' : entry.ppm, entry.kind === 'off' ? '' : entry.target,
        entry.stage, str_(b.water, 30), entry.summary, estado]);
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

    // Desde acá, solo con la contraseña de ajustes
    if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });

    if (b.action === 'setconfig') {
      if (!b.config || typeof b.config !== 'object') return out_({ ok: false, error: 'config' });
      setJSON_('CONFIG', b.config);
      return out_({ ok: true });
    }

    if (b.action === 'list') {
      const u = users_();
      return out_({
        ok: true,
        users: u.data.slice(1).map((r) => ({ id: r[0], name: r[1], device: r[2], first: r[3], last: r[4], count: r[5] })),
      });
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
