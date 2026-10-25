/**
 * Заявки на объекте — серверная часть (Google Apps Script, привязан к Google-таблице).
 * Приложение: https://kysylya88-bot.github.io/zayavki-app/
 *
 * Что делает:
 *  - вход сотрудников по рабочей почте (@upiterstroy.ru) и паролю, пароли выдаются письмом;
 *  - принимает заявки с телефонов (очередь без интернета) и хранит их в таблице;
 *  - считает остатки с учётом заявок ВСЕХ сотрудников, чтобы не было дублей;
 *  - отдаёт приложению реестр объекта (лист «Позиции»);
 *  - меню «Заявки» для снабжения: выдать пароли, загрузить реестр, выгрузить csv для Excel-реестра.
 *
 * Установка: см. лист «Инструкция» в таблице.
 */

var DOMAIN = '@upiterstroy.ru';
var DEFAULT_APP_URL = 'https://kysylya88-bot.github.io/zayavki-app/';
var TOKEN_DAYS = 90;
var MAX_FAILS = 5;
var PRIV_ROLES = ['Директор', 'Снабжение', 'Админ'];

var SH = { USERS: 'Пользователи', OBJ: 'Объекты', ITEMS: 'Позиции', REQS: 'Заявки', LINES: 'Строки', LOG: 'Журнал' };
var HDR = {
  USERS: ['Email', 'ФИО', 'Роль', 'Объекты', 'Активен', 'Соль', 'Хэш', 'Пароль выдан', 'Версия'],
  OBJ: ['Код', 'Название', 'Метка', 'Дата реестра', 'Email снабжения', 'Раздел Bitrix', 'След. номер (реестр)', 'Версии разделов', 'Активен'],
  ITEMS: ['Объект', 'Раздел', 'Код', 'Группа', 'Наименование', 'Ед.', 'По проекту', 'Заказано (реестр)', 'Осталось (реестр)', 'Статус'],
  REQS: ['Объект', '№', 'Дата', 'Автор (email)', 'Автор', 'К дате', 'Статус', 'Корректирует №', 'Версия проекта', 'В реестре', 'Позиций', 'Обновлено', 'cid', 'Предупреждения', 'Правок'],
  LINES: ['Объект', '№ заявки', 'Раздел', 'Код', 'Название', 'Новое название', 'Ед.', 'Кол-во', 'Аналог', 'Примечание', 'Заменяет (код)', 'Заменяет (название)', 'Кол-во замены', 'К дате', 'cid'],
  LOG: ['Время', 'Email', 'Действие', 'Подробности']
};
// столбцы с датами/временем хранятся как текст
var TEXT_COLS = { OBJ: [4], REQS: [3, 6, 12], LINES: [14], USERS: [8], LOG: [1] };

/* ===================== служебное ===================== */

function ss_() { return SpreadsheetApp.getActive(); }

function sheet_(k) {
  var s = ss_().getSheetByName(SH[k]);
  if (!s) {
    s = ss_().insertSheet(SH[k]);
    s.appendRow(HDR[k]);
    var tc = TEXT_COLS[k] || [];
    for (var i = 0; i < tc.length; i++) s.getRange(2, tc[i], 2000, 1).setNumberFormat('@');
  }
  return s;
}

function readAll_(k) {
  var s = sheet_(k), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, HDR[k].length).getValues();
}

function writeAll_(k, rows) {
  var s = sheet_(k), n = HDR[k].length, last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, n).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, n).setValues(rows);
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Europe/Moscow', 'yyyy-MM-dd');
  return String(v).trim();
}
function num_(v) { var n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function yes_(v) { var s = str_(v).toLowerCase(); return s === 'да' || s === 'true' || s === '1' || s === 'yes'; }
function no_(v) { var s = str_(v).toLowerCase(); return s === 'нет' || s === 'false' || s === '0' || s === 'no'; }
function now_() { return Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd HH:mm'); }
function hex_(bytes) {
  var o = '';
  for (var i = 0; i < bytes.length; i++) { var b = (bytes[i] + 256) % 256; o += (b < 16 ? '0' : '') + b.toString(16); }
  return o;
}
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }

function log_(email, action, details) {
  try { sheet_('LOG').appendRow([now_(), email || '', action, String(details || '').slice(0, 500)]); } catch (e) {}
}

function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function appUrl_() { return PropertiesService.getScriptProperties().getProperty('APP_URL') || DEFAULT_APP_URL; }

function randPass_() {
  var chars = 'abcdefghjkmnpqrstuvwxyzACDEFGHJKLMNPQRSTUVWXYZ23456789', o = '';
  for (var i = 0; i < 10; i++) o += chars.charAt(Math.floor(Math.random() * chars.length));
  return o;
}

/* ===================== токены и вход ===================== */

function makeToken_(email, ver) {
  var payload = JSON.stringify({ e: email, x: Date.now() + TOKEN_DAYS * 86400000, v: String(ver || '1') });
  var b = Utilities.base64EncodeWebSafe(payload);
  var sig = hex_(Utilities.computeHmacSha256Signature(b, secret_()));
  return b + '.' + sig;
}

function findUser_(email) {
  var rows = readAll_('USERS');
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][0]).toLowerCase() === email) return { i: i, row: rows[i] };
  }
  return null;
}

function userObjects_(row) {
  var all = readAll_('OBJ'), want = str_(row[3]), res = [], i;
  var codes = want.toLowerCase() === 'все' || want === '' ? null : want.split(/[,;\s]+/).filter(function (x) { return x; });
  for (i = 0; i < all.length; i++) {
    var code = str_(all[i][0]);
    if (!code || no_(all[i][8])) continue;
    if (codes === null || codes.indexOf(code) >= 0) res.push({ code: code, label: str_(all[i][2]) || str_(all[i][1]) || code });
  }
  return res;
}

function userInfo_(row) {
  return { email: str_(row[0]).toLowerCase(), name: str_(row[1]), role: str_(row[2]) || 'Прораб', ver: str_(row[8]) || '1' };
}

function auth_(token) {
  if (!token || String(token).indexOf('.') < 0) throw new Error('auth');
  var parts = String(token).split('.');
  var good = hex_(Utilities.computeHmacSha256Signature(parts[0], secret_()));
  if (good !== parts[1]) throw new Error('auth');
  var p;
  try { p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); } catch (e) { throw new Error('auth'); }
  if (!p.e || p.x < Date.now()) throw new Error('auth');
  var u = findUser_(p.e);
  if (!u || !yes_(u.row[4])) throw new Error('auth');
  var info = userInfo_(u.row);
  if (info.ver !== p.v) throw new Error('auth');
  info.objects = userObjects_(u.row);
  return info;
}

// нормализация ФИО: нижний регистр, ё=е, без точек и лишних пробелов
function normName_(v) { return str_(v).toLowerCase().replace(/ё/g, 'е').replace(/[.,]/g, ' ').replace(/\s+/g, ' ').trim(); }

// Вход по фамилии (без пароля). Можно написать «Иванов» или «Иванов И.П.».
// Если однофамильцев несколько, нужны инициалы.
function login_(o) {
  var q = normName_(o.name);
  if (!q) return { ok: false, error: 'Введите фамилию' };
  var cache = CacheService.getScriptCache(), fk = 'f:' + q;
  var fails = parseInt(cache.get(fk) || '0', 10);
  if (fails >= MAX_FAILS) return { ok: false, error: 'Слишком много попыток. Попробуйте через 15 минут.' };
  var qt = q.split(' '), qs = qt[0], qi = qt.slice(1).map(function (t) { return t.charAt(0); }).join('');
  if (qt.length === 2 && qt[1].length > 1 && qt[1].length <= 3) qi = qt[1]; // «И.П.» без пробела -> «ип»
  var rows = readAll_('USERS'), hits = [], i;
  for (i = 0; i < rows.length; i++) {
    if (!yes_(rows[i][4])) continue;
    var ut = normName_(rows[i][1]).split(' ');
    if (ut[0] !== qs) continue;
    var ui = ut.slice(1).map(function (t) { return t.charAt(0); }).join('');
    if (qi === '' || ui.indexOf(qi) === 0) hits.push(i);
  }
  if (hits.length === 0) {
    cache.put(fk, String(fails + 1), 900);
    log_(q, 'вход: не найден', '');
    return { ok: false, error: 'Такой фамилии нет в списке. Проверьте написание или обратитесь в снабжение.' };
  }
  if (hits.length > 1) return { ok: false, error: 'Есть однофамильцы. Введите фамилию и инициалы, например «' + str_(rows[hits[0]][1]).split(' ')[0] + ' И.О.»' };
  var row = rows[hits[0]], email = str_(row[0]).toLowerCase();
  cache.remove(fk);
  var info = userInfo_(row);
  log_(email, 'вход по фамилии', '');
  return { ok: true, token: makeToken_(email, info.ver), name: info.name, role: info.role, email: email, objects: userObjects_(row) };
}

function passwd_(o) {
  var user = auth_(o.token), neu = String(o.neu || '');
  if (neu.length < 8) return { ok: false, error: 'Пароль должен быть не короче 8 символов' };
  var u = findUser_(user.email);
  if (sha_(str_(u.row[5]) + String(o.old || '')) !== str_(u.row[6])) return { ok: false, error: 'Старый пароль неверный' };
  var salt = Utilities.getUuid(), s = sheet_('USERS');
  s.getRange(u.i + 2, 6).setValue(salt);
  s.getRange(u.i + 2, 7).setValue(sha_(salt + neu));
  log_(user.email, 'смена пароля', '');
  return { ok: true };
}

/* ===================== заявки ===================== */

function objRow_(code) {
  var all = readAll_('OBJ');
  for (var i = 0; i < all.length; i++) if (str_(all[i][0]) === code) return { i: i, row: all[i] };
  return null;
}

function parseLines_(lines) {
  var out = [];
  for (var i = 0; i < (lines || []).length; i++) {
    var l = lines[i];
    out.push({
      sect: str_(l.sect), code: str_(l.code), name: str_(l.name), newName: str_(l.newName), unit: str_(l.unit),
      qty: num_(l.qty), analog: l.analog ? 1 : 0, note: str_(l.note).slice(0, 300), replCode: str_(l.replCode),
      replName: str_(l.replName), replQty: num_(l.replQty), due: str_(l.due)
    });
  }
  return out;
}

// карта заявок объекта: cid -> {meta, lines[]}
function loadReqs_(obj) {
  var reqs = readAll_('REQS'), lines = readAll_('LINES'), m = {}, order = [], i;
  for (i = 0; i < reqs.length; i++) {
    var r = reqs[i];
    if (str_(r[0]) !== obj) continue;
    var cid = str_(r[12]);
    m[cid] = { row: i, no: num_(r[1]), date: str_(r[2]), email: str_(r[3]).toLowerCase(), author: str_(r[4]), due: str_(r[5]), status: str_(r[6]),
      corr: str_(r[7]), ver: str_(r[8]), reg: yes_(r[9]), upd: str_(r[11]), cid: cid, warn: str_(r[13]), rev: num_(r[14]), lines: [] };
    order.push(cid);
  }
  for (i = 0; i < lines.length; i++) {
    var l = lines[i], c = str_(l[14]);
    if (m[c]) m[c].lines.push({
      sect: str_(l[2]), code: str_(l[3]), name: str_(l[4]), newName: str_(l[5]), unit: str_(l[6]), qty: num_(l[7]), analog: yes_(l[8]) || str_(l[8]) === '1',
      note: str_(l[9]), replCode: str_(l[10]), replName: str_(l[11]), replQty: num_(l[12]), due: str_(l[13])
    });
  }
  return { map: m, order: order };
}

function active_(r) { return !r.reg && r.status !== 'Отменена'; }

// сумма по коду в незагруженных заявках, кроме заявок с cid из exceptCid и (опционально) автора
function pending_(data, exceptCid, exceptEmail) {
  var p = {};
  for (var k = 0; k < data.order.length; k++) {
    var r = data.map[data.order[k]];
    if (!active_(r)) continue;
    if (exceptCid && r.cid === exceptCid) continue;
    if (exceptEmail && r.email === exceptEmail) continue;
    for (var j = 0; j < r.lines.length; j++) {
      var l = r.lines[j];
      if (l.code) p[l.code] = (p[l.code] || 0) + l.qty;
      if (l.replCode) p[l.replCode] = (p[l.replCode] || 0) + l.replQty;
    }
  }
  return p;
}

function itemsRemaining_(obj) {
  var rows = readAll_('ITEMS'), m = {};
  for (var i = 0; i < rows.length; i++) if (str_(rows[i][0]) === obj) m[str_(rows[i][2])] = num_(rows[i][8]);
  return m;
}

function nextNo_(obj, data, objRow) {
  var n = objRow ? Math.max(1, num_(objRow.row[6])) : 1;
  for (var k in data.map) if (data.map[k].no >= n) n = data.map[k].no + 1;
  return n;
}

function sameLines_(a, b) {
  return JSON.stringify(parseLines_(a)) === JSON.stringify(parseLines_(b));
}

function upsert_(obj, user, r, data, objRow) {
  var cid = str_(r.cid);
  if (!cid) return { cid: '', error: 'нет идентификатора заявки' };
  var lines = parseLines_(r.lines);
  if (!lines.length) return { cid: cid, error: 'пустая заявка' };
  var ex = data.map[cid], priv = user.role === 'Админ';
  var status = r.sent ? 'Отправлена в Bitrix' : 'Сформирована';
  var no, rev = num_(r.rev);
  if (ex) {
    if (ex.email !== user.email && !priv) return { cid: cid, error: 'чужая заявка' };
    if (ex.reg) return { cid: cid, no: ex.no, rejected: 'Заявка уже загружена в реестр — правка закрыта', reg: true };
    if (ex.status === 'Отменена') return { cid: cid, no: ex.no, rejected: 'Заявка отменена снабжением' };
    if (ex.status === 'Отправлена в Bitrix' && !r.sent) {
      if (!sameLines_(ex.lines, lines)) return { cid: cid, no: ex.no, rejected: 'Заявка уже отправлена — выпустите корректирующую', sent: true };
    }
    if (ex.status === 'Отправлена в Bitrix' && r.sent && !sameLines_(ex.lines, lines)) {
      return { cid: cid, no: ex.no, rejected: 'Заявка уже отправлена — выпустите корректирующую', sent: true };
    }
    no = ex.no;
  } else {
    no = num_(r.no);
    var taken = false;
    for (var k in data.map) if (data.map[k].no === no && k !== cid) taken = true;
    if (!no || taken) no = nextNo_(obj, data, objRow);
  }
  // остатки: считаем с учётом этой заявки
  var tmp = { map: {}, order: [] }, kk;
  for (kk in data.map) { if (kk !== cid) { tmp.map[kk] = data.map[kk]; tmp.order.push(kk); } }
  tmp.map[cid] = { cid: cid, reg: false, status: status, email: user.email, lines: lines };
  tmp.order.push(cid);
  var pend = pending_(tmp), rem = itemsRemaining_(obj), warns = [], seen = {};
  for (var i = 0; i < lines.length; i++) {
    var codes = [lines[i].code, lines[i].replCode];
    for (var j = 0; j < codes.length; j++) {
      var c = codes[j];
      if (!c || seen[c] || rem[c] === undefined) continue;
      seen[c] = 1;
      var over = (pend[c] || 0) - rem[c];
      if (over > 1e-9) warns.push(c + ': заказано больше остатка на ' + (Math.round(over * 100) / 100));
    }
  }
  var warn = warns.join('; ');
  var vals = [obj, no, str_(r.date), user.email, user.name, str_(r.due), status, str_(r.corr), str_(r.ver),
    ex && ex.reg ? 'Да' : '', lines.length, now_(), cid, warn, rev];
  var reqSheet = sheet_('REQS');
  if (ex) {
    reqSheet.getRange(ex.row + 2, 1, 1, vals.length).setValues([vals]);
    // порядок хранения строк в data остаётся прежним
    ex.no = no; ex.status = status; ex.lines = lines; ex.due = str_(r.due);
  } else {
    reqSheet.appendRow(vals);
    data.map[cid] = { row: -1, no: no, date: str_(r.date), email: user.email, author: user.name, due: str_(r.due), status: status, corr: str_(r.corr),
      ver: str_(r.ver), reg: false, cid: cid, warn: warn, rev: rev, lines: lines };
    data.order.push(cid);
    // позиция новой строки в листе — конец; пересчитаем индекс
    data.map[cid].row = reqSheet.getLastRow() - 2;
  }
  // строки заявки: удалить старые, добавить новые
  var allLines = readAll_('LINES'), keep = [], x;
  for (x = 0; x < allLines.length; x++) if (str_(allLines[x][14]) !== cid) keep.push(allLines[x]);
  for (x = 0; x < lines.length; x++) {
    var l = lines[x];
    keep.push([obj, no, l.sect, l.code, l.name, l.newName, l.unit, l.qty, l.analog ? 'Да' : '', l.note, l.replCode, l.replName, l.replQty || '', l.due || str_(r.due), cid]);
  }
  writeAll_('LINES', keep);
  data.map[cid].lines = lines;
  return { cid: cid, no: no, warn: warn, status: status };
}

function snapshotPayload_(objRow, obj) {
  var o = objRow.row, rows = readAll_('ITEMS'), items = [], i, ver = {};
  try { ver = JSON.parse(str_(o[7]) || '{}'); } catch (e) { ver = {}; }
  for (i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (str_(r[0]) !== obj) continue;
    items.push([str_(r[1]), str_(r[2]), str_(r[3]), str_(r[4]), str_(r[5]), num_(r[6]), num_(r[7]), num_(r[8]), str_(r[9])]);
  }
  return { v: 1, object: str_(o[1]), label: str_(o[2]) || str_(o[1]), email: str_(o[4]), parent: str_(o[5]), date: str_(o[3]),
    nextReq: num_(o[6]) || 1, ver: ver, items: items };
}

function sync_(o) {
  var user = auth_(o.token), obj = str_(o.obj), allowed = false, i;
  for (i = 0; i < user.objects.length; i++) if (user.objects[i].code === obj) allowed = true;
  if (!allowed) return { ok: false, error: 'Нет доступа к объекту ' + obj };
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var objRow = objRow_(obj);
    if (!objRow) return { ok: false, error: 'Объект не найден' };
    var data = loadReqs_(obj), results = [];
    var out = o.outbox || [];
    for (i = 0; i < out.length; i++) {
      try { results.push(upsert_(obj, user, out[i], data, objRow)); }
      catch (e) { results.push({ cid: str_(out[i] && out[i].cid), error: String(e.message || e) }); }
    }
    if (out.length) log_(user.email, 'заявки', obj + ': принято ' + out.length);
    var priv = PRIV_ROLES.indexOf(user.role) >= 0, list = [], maxNo = 0;
    for (var k = 0; k < data.order.length; k++) {
      var r = data.map[data.order[k]];
      if (r.no > maxNo) maxNo = r.no;
      var mine = r.email === user.email;
      if (!priv && !mine) continue;
      list.push({ cid: r.cid, no: r.no, date: r.date, due: r.due, author: r.author, email: r.email, status: r.status, corr: r.corr, ver: r.ver,
        reg: r.reg, sent: r.status === 'Отправлена в Bitrix', canceled: r.status === 'Отменена', warn: r.warn, rev: r.rev, mine: mine, lines: r.lines });
    }
    var res = { ok: true, user: { name: user.name, role: user.role, email: user.email }, objects: user.objects, results: results,
      pend: pending_(data, '', user.email), maxNo: maxNo, reqs: list, now: now_() };
    if (str_(o.snapVer) !== str_(objRow.row[3]) || o.forceSnap) res.snap = snapshotPayload_(objRow, obj);
    return res;
  } finally { lock.releaseLock(); }
}

/* ===================== веб-API ===================== */

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var o = JSON.parse(e.postData.contents);
    if (o.a === 'login') return reply_(login_(o));
    if (o.a === 'sync') return reply_(sync_(o));
    if (o.a === 'passwd') return reply_(passwd_(o));
    if (o.a === 'ping') return reply_({ ok: true });
    return reply_({ ok: false, error: 'unknown action' });
  } catch (err) {
    var m = String(err && err.message || err);
    return reply_({ ok: false, error: m === 'auth' ? 'auth' : 'Ошибка сервера: ' + m, auth: m === 'auth' });
  }
}

function doGet() { return ContentService.createTextOutput('Заявки: сервер работает').setMimeType(ContentService.MimeType.TEXT); }

/* ===================== меню администратора ===================== */

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Заявки')
      .addItem('Загрузить реестр объекта (файл из Excel)…', 'importSnapshotDialog')
      .addItem('Отметить выбранные заявки как загруженные в реестр', 'markLoaded')
      .addItem('Выгрузить заявки в csv для Excel-реестра', 'exportCsv')
      .addSeparator()
      .addItem('Первоначальная настройка', 'setup')
      .addToUi();
  } catch (e) {}
}

function setup() {
  for (var k in SH) sheet_(k);
  secret_();
  var ui = SpreadsheetApp.getUi();
  ui.alert('Готово', 'Листы проверены. Дальше: Развернуть → Новое развёртывание → Веб-приложение (доступ: все), и вставьте адрес в файл config.json приложения. Подробности на листе «Инструкция».', ui.ButtonSet.OK);
}

function sendPassword_(email, name, pass) {
  var body = 'Здравствуйте' + (name ? ', ' + name : '') + '!\n\n' +
    'Вам открыт доступ к приложению «Заявки на объекте».\n\n' +
    'Ссылка: ' + appUrl_() + '\n' +
    'Логин: ' + email + '\n' +
    'Пароль: ' + pass + '\n\n' +
    'Откройте ссылку на телефоне, введите логин и пароль. Потом добавьте приложение на главный экран:\n' +
    'iPhone (Safari): «Поделиться» → «На экран Домой». Android (Chrome): меню ⋮ → «Установить приложение».\n' +
    'Пароль можно сменить в меню приложения. Никому его не передавайте.';
  MailApp.sendEmail({ to: email, subject: 'Доступ к приложению «Заявки на объекте»', body: body });
}

function issueFor_(i, force) {
  var s = sheet_('USERS'), rows = readAll_('USERS'), r = rows[i];
  var email = str_(r[0]).toLowerCase();
  if (email.slice(-DOMAIN.length) !== DOMAIN) return 'пропущен (не ' + DOMAIN + '): ' + email;
  if (!yes_(r[4])) return 'пропущен (не активен): ' + email;
  if (str_(r[6]) && !force) return null;
  var pass = randPass_(), salt = Utilities.getUuid();
  s.getRange(i + 2, 6).setValue(salt);
  s.getRange(i + 2, 7).setValue(sha_(salt + pass));
  s.getRange(i + 2, 8).setValue(now_());
  s.getRange(i + 2, 9).setValue(String((parseInt(str_(r[8]) || '1', 10) || 1) + (force ? 1 : 0)));
  sendPassword_(email, str_(r[1]), pass);
  log_(email, 'пароль выдан', force ? 'сброс' : 'новый');
  return 'отправлено: ' + email;
}

function issuePasswords() {
  var ui = SpreadsheetApp.getUi(), rows = readAll_('USERS'), msgs = [], n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!str_(rows[i][0])) continue;
    try { var m = issueFor_(i, false); if (m) { msgs.push(m); if (m.indexOf('отправлено') === 0) n++; } }
    catch (e) { msgs.push('ошибка ' + str_(rows[i][0]) + ': ' + e.message); }
  }
  ui.alert('Пароли', 'Отправлено писем: ' + n + (msgs.length ? '\n\n' + msgs.join('\n') : '\nНовых сотрудников без пароля нет.'), ui.ButtonSet.OK);
}

function resetPassword() {
  var ui = SpreadsheetApp.getUi(), sh = ss_().getActiveSheet();
  if (sh.getName() !== SH.USERS) { ui.alert('Откройте лист «Пользователи» и выделите строку сотрудника.'); return; }
  var row = ss_().getActiveRange().getRow();
  if (row < 2) { ui.alert('Выделите строку сотрудника (не заголовок).'); return; }
  try { ui.alert(issueFor_(row - 2, true) || 'Не выполнено'); } catch (e) { ui.alert('Ошибка: ' + e.message); }
}

function importSnapshotDialog() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px">' +
    '<p>Файл <b>Снимок_реестра_….json</b> выгружается из Excel-реестра кнопкой «Снимок реестра для приложения».</p>' +
    '<p>Код объекта (латиницей, без пробелов, например B55):<br><input id="c" style="width:200px" value="B55"></p>' +
    '<p><input type="file" id="f" accept=".json"></p>' +
    '<p><button onclick="go()">Загрузить</button> <span id="m"></span></p></div>' +
    '<script>function go(){var f=document.getElementById("f").files[0],c=document.getElementById("c").value.trim();' +
    'if(!f||!c){document.getElementById("m").textContent="Выберите файл и укажите код";return}' +
    'var r=new FileReader();r.onload=function(){document.getElementById("m").textContent="Загружаю…";' +
    'google.script.run.withSuccessHandler(function(t){document.getElementById("m").textContent=t}).withFailureHandler(function(e){document.getElementById("m").textContent="Ошибка: "+e.message}).importSnapshot(c,r.result)};r.readAsText(f,"utf-8")}</script>'
  ).setWidth(520).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Загрузить реестр объекта');
}

function importSnapshot(code, text) {
  code = str_(code);
  if (!/^[A-Za-z0-9_-]+$/.test(code)) throw new Error('Код объекта: только латиница, цифры, - и _');
  var o = JSON.parse(String(text).replace(/^﻿/, ''));
  if (!o.items || !o.v) throw new Error('Это не файл снимка реестра');
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var objs = readAll_('OBJ'), found = -1, i;
    for (i = 0; i < objs.length; i++) if (str_(objs[i][0]) === code) found = i;
    var row = [code, str_(o.object), str_(o.label), str_(o.date), str_(o.email), str_(o.parent), num_(o.nextReq) || 1, JSON.stringify(o.ver || {}), found >= 0 ? objs[found][8] : 'Да'];
    if (found >= 0) objs[found] = row; else objs.push(row);
    writeAll_('OBJ', objs);
    var items = readAll_('ITEMS'), keep = [];
    for (i = 0; i < items.length; i++) if (str_(items[i][0]) !== code) keep.push(items[i]);
    for (i = 0; i < o.items.length; i++) {
      var it = o.items[i];
      keep.push([code, it[0], it[1], it[2], it[3], it[4], it[5], it[6], it[7], it[8]]);
    }
    writeAll_('ITEMS', keep);
  } finally { lock.releaseLock(); }
  log_(Session.getActiveUser().getEmail(), 'реестр загружен', code + ': ' + o.items.length + ' поз., от ' + o.date);
  return 'Готово: ' + o.items.length + ' позиций, реестр от ' + o.date + '. Не забудьте отметить загруженные в реестр заявки.';
}

function selectedReqRows_() {
  var sh = ss_().getActiveSheet();
  if (sh.getName() !== SH.REQS) return null;
  var rg = ss_().getActiveRange(), res = [], r0 = rg.getRow(), n = rg.getNumRows();
  for (var r = r0; r < r0 + n; r++) if (r >= 2) res.push(r - 2);
  return res;
}

function markLoaded() {
  var ui = SpreadsheetApp.getUi(), sel = selectedReqRows_();
  if (!sel || !sel.length) { ui.alert('Откройте лист «Заявки» и выделите строки заявок, которые уже загружены в Excel-реестр.'); return; }
  var s = sheet_('REQS'), rows = readAll_('REQS'), n = 0;
  for (var i = 0; i < sel.length; i++) {
    if (sel[i] < rows.length && str_(rows[sel[i]][12])) { s.getRange(sel[i] + 2, 10).setValue('Да'); n++; }
  }
  log_(Session.getActiveUser().getEmail(), 'в реестре', 'отмечено заявок: ' + n);
  ui.alert('Отмечено заявок: ' + n + '. Они больше не уменьшают остатки (теперь остаток берётся из реестра), правка этих заявок в приложении закрыта.');
}

function csvClean_(s) { return String(s === null || s === undefined ? '' : s).replace(/[;\r\n\t]+/g, ' ').trim(); }

function buildCsv_(rows) {
  var reqs = readAll_('REQS'), lines = readAll_('LINES'), want = {}, i;
  for (i = 0; i < rows.length; i++) want[str_(reqs[rows[i]][12])] = reqs[rows[i]];
  var objs = readAll_('OBJ'), labelOf = {};
  for (i = 0; i < objs.length; i++) labelOf[str_(objs[i][0])] = str_(objs[i][2]) || str_(objs[i][1]);
  var first = want[Object.keys(want)[0]];
  var out = ['BITREQ1;' + csvClean_(labelOf[str_(first[0])] || '') + ';' + str_(first[2]) + ';' + csvClean_(first[4])];
  for (i = 0; i < lines.length; i++) {
    var l = lines[i], c = str_(l[14]);
    if (!want[c]) continue;
    var rq = want[c], rc = str_(l[10]);
    out.push(['L', num_(rq[1]), str_(rq[2]), str_(l[2]), str_(l[3]), csvClean_(l[4]), csvClean_(l[5]), num_(l[7]), yes_(l[8]) ? 1 : 0, csvClean_(l[9]),
      rc ? csvClean_(rc + ' | ' + str_(l[11])) : '', str_(l[13]) || str_(rq[5]), csvClean_(l[6]), num_(l[12]), csvClean_(rq[4])].join(';'));
  }
  return '﻿' + out.join('\r\n');
}

function exportCsv() {
  var ui = SpreadsheetApp.getUi(), sel = selectedReqRows_(), reqs = readAll_('REQS'), rows = [], i;
  if (sel && sel.length) { for (i = 0; i < sel.length; i++) if (sel[i] < reqs.length && str_(reqs[sel[i]][12])) rows.push(sel[i]); }
  else { for (i = 0; i < reqs.length; i++) if (str_(reqs[i][12]) && !yes_(reqs[i][9]) && str_(reqs[i][6]) !== 'Отменена') rows.push(i); }
  if (!rows.length) { ui.alert('Нет заявок для выгрузки. (Выделите строки на листе «Заявки» или оставьте выделение пустым, чтобы выгрузить все не загруженные в реестр.)'); return; }
  var f = DriveApp.createFile('Заявки_для_реестра_' + now_().replace(/[: ]/g, '-') + '.csv', buildCsv_(rows), MimeType.CSV);
  log_(Session.getActiveUser().getEmail(), 'выгрузка csv', 'заявок: ' + rows.length);
  ui.alert('Файл создан в вашем Google Диске:\n' + f.getName() + '\n' + f.getUrl() + '\n\nСкачайте его и загрузите в Excel-реестр кнопкой «Загрузить заявки из приложения (csv)». Потом отметьте эти заявки как загруженные.');
}
