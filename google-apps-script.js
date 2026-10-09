/**
 * СКРИПТ ДЛЯ ХРАНЕНИЯ КЛИЕНТОВ CRM «СТРОЙ ИНДАСТРИС» В GOOGLE ТАБЛИЦЕ
 * ===================================================================
 *
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 *
 * 1. Создайте НОВУЮ Google Таблицу (sheets.google.com). Не используйте таблицу
 *    от CRM «Этно Консалт»: скрипт сам создаст в ней лист «Клиенты» с нужными
 *    колонками.
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
 *
 * 8. Откройте CRM (index.html), нажмите кнопку "Google Таблица" над таблицей
 *    клиентов, вставьте URL и нажмите "Подключить".
 *
 * 9. Готово. Каждый клиент — одна строка на листе «Клиенты». При изменении
 *    клиента в CRM его строка обновляется, при удалении — удаляется.
 *
 * ВАЖНО:
 * - Этот URL даёт доступ ко всей базе клиентов. Не публикуйте его и не
 *   вставляйте в код сайта — CRM хранит его только в вашем браузере.
 * - Колонку "ID" в таблице не меняйте и не удаляйте: по ней CRM находит строку
 *   клиента. Остальные ячейки можно править руками — CRM подхватит изменения
 *   при следующей синхронизации.
 * - Если вы позже измените код скрипта, нужно сделать НОВОЕ развёртывание
 *   версии (Deploy → Manage deployments → Edit → New version), иначе изменения
 *   не применятся к уже выданному URL.
 */

var SHEET_NAME = 'Клиенты';

// Порядок колонок в таблице: ключ поля в CRM → заголовок
var COLUMNS = [
  ['id', 'ID'],
  ['createdAt', 'Добавлен'],
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

function doGet() {
  return json({ result: 'success', message: 'CRM Строй Индастрис: скрипт работает.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    if (body.action === 'list') {
      return json({ result: 'success', clients: readClients(sheet) });
    }
    if (body.action === 'save') {
      (body.deletes || []).forEach(function (id) { deleteClient(sheet, id); });
      (body.clients || []).forEach(function (client) { upsertClient(sheet, client); });
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
  if (DATE_KEYS.indexOf(key) !== -1) return value.trim().replace(' ', 'T');
  return value;
}

function upsertClient(sheet, client) {
  if (!client || !client.id) return;
  var row = COLUMNS.map(function (c) { return toCell(c[0], client); });
  var index = findRow(sheet, client.id) || sheet.getLastRow() + 1;
  sheet.getRange(index, 1, 1, COLUMNS.length).setNumberFormat('@').setValues([row]);
}

function deleteClient(sheet, id) {
  var index = findRow(sheet, id);
  if (index) sheet.deleteRow(index);
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
