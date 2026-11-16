// Registro de usuarios de Dosis Hidroponia.
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

function hash_(pin) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'hidro:' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}

// La primera vez que abrís "Usuarios" en la app, la contraseña de ajustes queda guardada acá.
// Si cambiás la contraseña de ajustes: Configuración del proyecto > Propiedades del script > borrar ADMIN.
function checkPin_(pin) {
  if (!pin) return false;
  const props = PropertiesService.getScriptProperties();
  const h = hash_(String(pin));
  const saved = props.getProperty('ADMIN');
  if (!saved) {
    props.setProperty('ADMIN', h);
    return true;
  }
  return saved === h;
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

    if (b.action === 'hit') {
      const id = String(b.id || '').slice(0, 40);
      if (!id) return out_({ ok: false, error: 'id' });
      const dev = String(b.device || '').slice(0, 60);
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

    if (!checkPin_(b.pin)) return out_({ ok: false, error: 'pin' });

    if (b.action === 'list') {
      return out_({
        ok: true,
        users: data.slice(1).map((r) => ({ id: r[0], name: r[1], device: r[2], first: r[3], last: r[4], count: r[5] })),
      });
    }

    if (b.action === 'rename') {
      const r = rowOf(String(b.id));
      if (!r) return out_({ ok: false, error: 'id' });
      users.getRange(r, 2).setValue(String(b.name || '').slice(0, 40));
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
