// Registro de Dosis Hidroponia: usuarios, mediciones de tanques y ajustes compartidos.
// Proyecto de Apps Script publicado como "Aplicación web".

// Planilla "Hidro Dosis Usuarios"
const SS = SpreadsheetApp.openById('1vxmWFBLpxiWxihmsyGTzwQ2NOoUXO5o54gotwPuVGYA');

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

function doGet() {
  return out_({ ok: true, app: 'hidro-dosis' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'ping') return out_({ ok: true });

    const users = sheet_('Usuarios', ['ID', 'Nombre', 'Dispositivo', 'Primer ingreso', 'Último ingreso', 'Ingresos']);
    const data = users.getDataRange().getValues();
    const rowOf = (id) => {
      for (let i = 1; i < data.length; i++) if (data[i][0] === id) return i + 1;
      return 0;
    };
    const userName = (id) => {
      const r = rowOf(String(id || ''));
      return r ? String(data[r - 1][1]) : 'Sin registrar';
    };

    if (b.action === 'hit') {
      const id = str_(b.id, 40);
      if (!id) return out_({ ok: false, error: 'id' });
      const dev = str_(b.device, 60);
      const now = new Date();
      const r = rowOf(id);
      let name;
      if (r) {
        name = data[r - 1][1];
        users.getRange(r, 3).setValue(dev);
        users.getRange(r, 5).setValue(now);
        users.getRange(r, 6).setValue((Number(data[r - 1][5]) || 0) + 1);
      } else {
        name = 'Dispositivo ' + data.length;
        users.appendRow([id, name, dev, now, now, 1]);
      }
      sheet_('Ingresos', ['Fecha', 'ID', 'Nombre', 'Dispositivo']).appendRow([now, id, name, dev]);
      return out_({ ok: true });
    }

    // Estado de todos los tanques y ajustes compartidos
    if (b.action === 'state') {
      return out_({ ok: true, config: getJSON_('CONFIG', null), last: getJSON_('LAST', {}) });
    }

    const log = () => sheet_('Mediciones', ['Fecha', 'Tanque', 'Usuario', 'Litros', 'ppm', 'Objetivo', 'Etapa', 'Llenar', 'Indicación', 'Estado']);

    // Nueva medición de un tanque
    if (b.action === 'measure') {
      const tank = str_(b.tank, 10);
      if (!tank) return out_({ ok: false, error: 'tank' });
      const now = new Date();
      const entry = {
        time: now.toISOString(),
        user: userName(b.id),
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
      log().appendRow([now, str_(b.name, 30) || tank, entry.user, entry.liters, entry.ppm, entry.target, entry.stage,
        entry.fill ? 'Sí' : 'No', entry.summary, entry.kind === 'ok' ? 'En rango' : 'Para corregir']);
      return out_({ ok: true, last: entry });
    }

    // Marcar como corregido
    if (b.action === 'done') {
      const tank = str_(b.tank, 10);
      const last = getJSON_('LAST', {});
      if (!last[tank]) return out_({ ok: false, error: 'tank' });
      const now = new Date();
      last[tank].done = { time: now.toISOString(), user: userName(b.id) };
      setJSON_('LAST', last);
      log().appendRow([now, str_(b.name, 30) || tank, last[tank].done.user, '', '', '', '', '', 'Marcado como corregido', 'Corregido']);
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
      return out_({
        ok: true,
        users: data.slice(1).map((r) => ({ id: r[0], name: r[1], device: r[2], first: r[3], last: r[4], count: r[5] })),
      });
    }

    if (b.action === 'rename') {
      const r = rowOf(String(b.id));
      if (!r) return out_({ ok: false, error: 'id' });
      users.getRange(r, 2).setValue(str_(b.name, 40));
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
