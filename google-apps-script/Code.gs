/**
 * Скан в Таблицу — серверная часть (Google Apps Script).
 *
 * Установка (один раз):
 *   1. Откройте нужную Google Таблицу → Расширения → Apps Script.
 *   2. Удалите всё содержимое Code.gs и вставьте этот файл целиком.
 *   3. (Необязательно) задайте SECRET ниже — тот же ключ укажите в настройках приложения.
 *   4. Развернуть → Новое развертывание → тип «Веб-приложение»:
 *        Выполнять как: Я
 *        У кого есть доступ: Все
 *   5. Скопируйте URL веб-приложения (заканчивается на /exec) в настройки приложения.
 *
 * После изменения этого кода: Развернуть → Управление развертываниями →
 * карандаш → Версия: «Новая версия» → Развернуть (URL не меняется).
 */

// Секретный ключ. Пусто = без проверки. Если задан — такой же нужно ввести в приложении.
var SECRET = '';

// Лист по умолчанию, если приложение не прислало своё имя листа.
var DEFAULT_SHEET = 'Сканы';

var HEADERS = ['Дата и время', 'Штрих-код', 'Количество', 'Тип кода', 'Устройство', 'Комментарий', 'ID скана'];
var COL_CODE = 2;
var COL_ID = 7;
var VERSION = 1;

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (SECRET && params.token !== SECRET) {
    return json_({ ok: false, error: 'Неверный секретный ключ' });
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return json_({ ok: true, app: 'scan-to-sheet', version: VERSION, spreadsheet: ss.getName() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Несколько телефонов могут отправлять одновременно — пишем строго по очереди.
    lock.waitLock(30000);

    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (SECRET && data.token !== SECRET) {
      return json_({ ok: false, error: 'Неверный секретный ключ' });
    }

    var rows = Array.isArray(data.rows) ? data.rows : [];
    var sheet = getSheet_(data.sheet);
    var known = existingIds_(sheet);
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    var saved = [];
    var duplicates = [];
    var values = [];
    rows.forEach(function (r) {
      if (!r || !r.id || r.code === undefined || r.code === null) return;
      var id = String(r.id);
      // Повторная отправка того же скана (например, после обрыва связи) не создаёт дубль.
      if (known[id]) {
        duplicates.push(id);
        return;
      }
      known[id] = true;
      var ts = new Date(r.ts);
      if (isNaN(ts.getTime())) ts = new Date();
      values.push([
        ts,
        String(r.code),
        Number(r.qty) || 1,
        String(r.format || ''),
        String(r.device || ''),
        String(r.note || ''),
        id
      ]);
      saved.push(id);
    });

    if (values.length) {
      var start = sheet.getLastRow() + 1;
      var range = sheet.getRange(start, 1, values.length, HEADERS.length);
      // Штрих-код и ID как текст: иначе EAN-13 превращается в 4,6E+12 и теряются ведущие нули.
      sheet.getRange(start, COL_CODE, values.length, 1).setNumberFormat('@');
      sheet.getRange(start, COL_ID, values.length, 1).setNumberFormat('@');
      sheet.getRange(start, 1, values.length, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
      range.setValues(values);
      SpreadsheetApp.flush();
    }

    return json_({ ok: true, saved: saved, duplicates: duplicates, timeZone: tz });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = String(name || '').trim() || DEFAULT_SHEET;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(1, COL_CODE, sheet.getMaxRows(), 1).setNumberFormat('@');
  }
  return sheet;
}

function existingIds_(sheet) {
  var map = {};
  var last = sheet.getLastRow();
  if (last < 2) return map;
  var ids = sheet.getRange(2, COL_ID, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0]) map[String(ids[i][0])] = true;
  }
  return map;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
