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

const ADMIN_ACTIONS = ['history', 'setconfig', 'list', 'rename', 'block', 'remove'];

function doGet() {
  return out_({ ok: true, app: 'hidro-dosis' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.action === 'ping') return out_({ ok: true });

  // Un dispositivo bloqueado no lee ni guarda nada (en las acciones de ajustes, b.id es el dispositivo a modificar)
  if (ADMIN_ACTIONS.indexOf(b.action) < 0 && isBlocked_(b.id)) return out_({ ok: false, error: 'blocked' });

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
    const vals = s.getRange(start, 1, n - start + 1, COLS).getValues();
    return out_({
      ok: true,
      rows: vals.reverse().map((r) => ({ time: r[0], tank: r[1], user: r[2], liters: r[3], ppm: r[4], target: r[5], stage: r[6], water: r[7], summary: r[8], status: r[9], ph: r[10], note: r[11] })),
    });
  }

  // Registro de un día (para todos): movimientos entre from y to
  if (b.action === 'day') {
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
