/**
 * СКРИПТ ДЛЯ ХРАНЕНИЯ КЛИЕНТОВ CRM «СТРОЙ ИНДАСТРИС» В GOOGLE ТАБЛИЦЕ
 * ===================================================================
 *
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 *
 * 1. Создайте НОВУЮ Google Таблицу (sheets.google.com). Не используйте таблицу
 *    от CRM «Этно Консалт»: скрипт сам создаст в ней листы «Клиенты» и
 *    «Пользователи» с нужными колонками.
 *
 * 2. В таблице откройте меню "Расширения" → "Apps Script".
 *
 * 3. Удалите весь код-заготовку в открывшемся редакторе и вставьте вместо него
 *    ВЕСЬ код из этого файла (ниже).
 *
 * 4. Нажмите "Сохранить" (значок дискеты), дайте проекту любое имя.
 *
 * 5. Нажмите "Развернуть" (Deploy) → "Новое развёртывание" (New deployment).
 *    - Тип развёртывания: "Веб-приложение" (Web app).
 *    - Описание: любое, например "CRM".
 *    - "Кто выполняет доступ" (Execute as): "Я" (свой аккаунт).
 *    - "У кого есть доступ" (Who has access): "Все" (Anyone).
 *    Нажмите "Развернуть".
 *
 * 6. Google попросит разрешить доступ — разрешите (это ваш собственный скрипт,
 *    он работает только с вашей таблицей).
 *
 * 7. Скопируйте выданный URL вида:
 *    https://script.google.com/macros/s/XXXXXXXXXXXXXXXX/exec
 *    и один раз откройте его в браузере. Должна появиться надпись
 *    «скрипт работает», а в таблице — листы «Клиенты» и «Пользователи».
 *
 * 8. Откройте CRM (index.html), нажмите "Подключить таблицу", вставьте URL.
 *
 * ПОЛЬЗОВАТЕЛИ И ПАРОЛИ (лист «Пользователи»):
 *
 * - Каждая строка — один сотрудник: Логин, Пароль, Имя, Роль.
 * - Первая строка создаётся сама: логин admin, роль «админ», пароль случайный —
 *   посмотрите его в колонке «Пароль» и войдите с ним в CRM.
 * - Чтобы добавить сотрудника, допишите строку: логин, пароль, имя и роль
 *   «менеджер». Чтобы закрыть доступ — удалите строку.
 * - Роль «менеджер» видит и меняет только своих клиентов. Роль «админ» видит
 *   всех клиентов и может передавать их между менеджерами.
 * - После первого входа пароль в таблице заменяется на шифр (sha256$...), так
 *   что подсмотреть его уже нельзя. Чтобы сменить пароль, просто впишите в
 *   ячейку новый поверх шифра.
 *
 * ВАЖНО:
 * - Колонки "ID" и "Менеджер" на листе «Клиенты» руками не меняйте: по ним CRM
 *   находит строку клиента и её владельца. Остальные ячейки можно править —
 *   CRM подхватит изменения при следующей синхронизации.
 * - Доступ к самой Google Таблице давайте только тем, кому можно видеть всю
 *   базу и менять пароли.
 * - Если вы позже измените код скрипта, нужно сделать НОВОЕ развёртывание
 *   версии (Deploy → Manage deployments → Edit → New version), иначе изменения
 *   не применятся к уже выданному URL.
 */

var SHEET_NAME = 'Клиенты';
var USERS_SHEET_NAME = 'Пользователи';

// Порядок колонок в таблице: ключ поля в CRM → заголовок
var COLUMNS = [
  ['id', 'ID'],
  ['createdAt', 'Добавлен'],
  ['owner', 'Менеджер'],
  ['name', 'Имя'],
  ['org', 'Организация'],
  ['urgency', 'Как срочно'],
  ['product', 'Товар'],
  ['qty', 'Кол-во'],
  ['phone', 'Номер'],
  ['email', 'Почта'],
  ['labels', 'Ярлыки'],
  ['remindAt', 'Напоминание'],
  ['remindText', 'Текст напоминания'],
  ['updatedAt', 'Обновлён']
];
var DATE_KEYS = ['createdAt', 'remindAt', 'updatedAt'];
var OWNER_COLUMN = 3;

var USER_COLUMNS = ['Логин', 'Пароль', 'Имя', 'Роль'];
var HASH_PREFIX = 'sha256$';
var TOKEN_DAYS = 30;
var MAX_LOGIN_FAILS = 10;

function doGet() {
  getSheet();
  getUsersSheet();
  return json({ result: 'success', message: 'CRM Строй Индастрис: скрипт работает.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'login') return json(login(body.login, body.password));

    var user = userFromToken(body.token);
    if (!user) return json({ result: 'auth', message: 'нужно войти заново' });

    var sheet = getSheet();
    if (body.action === 'list') {
      var clients = readClients(sheet).filter(function (c) {
        return user.admin || c.owner === user.login;
      });
      var reply = { result: 'success', clients: clients };
      if (user.admin) reply.users = readUsers(getUsersSheet()).map(publicUser);
      return json(reply);
    }
    if (body.action === 'save') {
      (body.deletes || []).forEach(function (id) { deleteClient(sheet, id, user); });
      (body.clients || []).forEach(function (client) { upsertClient(sheet, client, user); });
      return json({ result: 'success' });
    }
    return json({ result: 'error', message: 'неизвестное действие' });
  } catch (err) {
    return json({ result: 'error', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- пользователи ----------

function getUsersSheet() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(USERS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, sheet.getMaxRows(), USER_COLUMNS.length).setNumberFormat('@');
    sheet.appendRow(USER_COLUMNS);
    sheet.getRange(1, 1, 1, USER_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Первый вход: пароль виден в таблице, пока им не воспользуются
    var password = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    sheet.appendRow(['admin', password, 'Руководитель', 'админ']);
  }
  return sheet;
}

function normLogin(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function readUsers(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, USER_COLUMNS.length).getDisplayValues();
  return rows.map(function (row, i) {
    return {
      row: i + 2,
      login: normLogin(row[0]),
      cell: row[1],
      name: row[2].trim() || row[0].trim(),
      admin: /^(админ|admin)/i.test(row[3].trim())
    };
  }).filter(function (u) { return u.login; });
}

function findUser(sheet, loginName) {
  loginName = normLogin(loginName);
  return readUsers(sheet).filter(function (u) { return u.login === loginName; })[0] || null;
}

function publicUser(user) {
  return { login: user.login, name: user.name, admin: user.admin };
}

function hashPassword(salt, password) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// Пароль, вписанный в таблицу руками, после первого входа заменяем на шифр
function checkPassword(sheet, user, password) {
  if (user.cell.indexOf(HASH_PREFIX) === 0) {
    var parts = user.cell.split('$');
    return parts.length === 3 && hashPassword(parts[1], password) === parts[2];
  }
  if (!user.cell || user.cell !== password) return false;
  var salt = Utilities.getUuid().replace(/-/g, '');
  user.cell = HASH_PREFIX + salt + '$' + hashPassword(salt, password);
  sheet.getRange(user.row, 2).setNumberFormat('@').setValue(user.cell);
  return true;
}

function login(loginName, password) {
  loginName = normLogin(loginName);
  password = String(password == null ? '' : password);
  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + loginName.slice(0, 100);
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_LOGIN_FAILS) {
    return { result: 'error', message: 'Слишком много неверных попыток. Подождите 10 минут.' };
  }
  var sheet = getUsersSheet();
  var user = findUser(sheet, loginName);
  if (!user || !password || !checkPassword(sheet, user, password)) {
    cache.put(failKey, String(fails + 1), 600);
    return { result: 'error', message: 'Неверный логин или пароль.' };
  }
  cache.remove(failKey);
  return { result: 'success', token: makeToken(user), user: publicUser(user) };
}

function getSecret() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('TOKEN_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', secret);
  }
  return secret;
}

function sign(text) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(text, getSecret(), Utilities.Charset.UTF_8));
}

// В подпись входит шифр пароля: смена пароля в таблице отменяет старые входы
function makeToken(user) {
  var expires = Date.now() + TOKEN_DAYS * 864e5;
  return Utilities.base64EncodeWebSafe(user.login, Utilities.Charset.UTF_8) + '.' + expires + '.' +
    sign(user.login + '|' + expires + '|' + user.cell);
}

function userFromToken(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 3 || !(Number(parts[1]) > Date.now())) return null;
  var loginName;
  try {
    loginName = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  } catch (err) {
    return null;
  }
  var user = findUser(getUsersSheet(), loginName);
  if (!user || user.cell.indexOf(HASH_PREFIX) !== 0) return null;
  return sign(user.login + '|' + parts[1] + '|' + user.cell) === parts[2] ? user : null;
}

// ---------- клиенты ----------

function getSheet() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    // Весь лист — текст, чтобы телефоны и даты не превращались в числа
    sheet.getRange(1, 1, sheet.getMaxRows(), COLUMNS.length).setNumberFormat('@');
    sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Номер строки клиента по ID или 0, если такого нет
function findRow(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ids = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === String(id)) return i + 2;
  }
  return 0;
}

function ownerOfRow(sheet, row) {
  return normLogin(sheet.getRange(row, OWNER_COLUMN).getDisplayValue());
}

function toCell(key, client) {
  var value = client[key];
  if (key === 'labels') value = (value || []).join(', ');
  value = String(value == null ? '' : value);
  if (DATE_KEYS.indexOf(key) !== -1) value = value.replace('T', ' ');
  // Не даём тексту из формы стать формулой
  if (/^[=+\-@]/.test(value)) value = "'" + value;
  return value;
}

function fromCell(key, value) {
  if (key === 'labels') {
    return value.split(',').map(function (s) { return s.trim(); }).filter(String);
  }
  if (key === 'owner') return normLogin(value);
  if (DATE_KEYS.indexOf(key) !== -1) return value.trim().replace(' ', 'T');
  return value;
}

// Менеджер пишет только в свои строки; владельца строки меняет только админ
function upsertClient(sheet, client, user) {
  if (!client || !client.id) return;
  var index = findRow(sheet, client.id);
  var current = index ? ownerOfRow(sheet, index) : '';
  if (index && !user.admin && current !== user.login) return;
  client.owner = user.admin ? (normLogin(client.owner) || current || user.login) : user.login;
  var row = COLUMNS.map(function (c) { return toCell(c[0], client); });
  index = index || sheet.getLastRow() + 1;
  sheet.getRange(index, 1, 1, COLUMNS.length).setNumberFormat('@').setValues([row]);
}

function deleteClient(sheet, id, user) {
  var index = findRow(sheet, id);
  if (!index) return;
  if (!user.admin && ownerOfRow(sheet, index) !== user.login) return;
  sheet.deleteRow(index);
}

function readClients(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, COLUMNS.length).getDisplayValues();
  return rows
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      var client = {};
      COLUMNS.forEach(function (c, i) { client[c[0]] = fromCell(c[0], row[i]); });
      return client;
    });
}
