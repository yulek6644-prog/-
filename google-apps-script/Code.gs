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
 *
 * Формат строки — такой же, как у Scan-IT to Office на листе «Данные»:
 *   A  артикул (штрих-код)
 *   B  дата (yyyy-mm-dd)
 *   C  наименование — формула ВПР по листу «номенклатура»
 *   D  количество (минус — расход, плюс — приход)
 *   E  комментарий (например, «П.907»)
 *   F  «-»
 *   G  устройство / сотрудник
 *   H  «-»
 *   I  время скана на телефоне
 *   J  время записи в таблицу
 *   K  ID скана — служебная колонка против дублей, не удаляйте и не меняйте
 */

// Секретный ключ. Пусто = без проверки. Если задан — такой же нужно ввести в приложении.
var SECRET = '';

// ID таблицы — нужен, только если скрипт создан отдельно на script.google.com, а не из меню таблицы.
// Берётся из адреса таблицы: docs.google.com/spreadsheets/d/ ЭТА_ЧАСТЬ /edit
// Пусто = таблица, из которой открыт редактор скриптов.
var SPREADSHEET_ID = '';

// Лист по умолчанию, если приложение не прислало своё имя листа.
var DEFAULT_SHEET = 'Данные';

// Лист со справочником: A — артикул, B — наименование.
var NOMENCLATURE_SHEET = 'номенклатура';

var HEADERS = ['Артикул', 'ДАТА', 'НАИМЕНОВАНИЕ', 'КОЛИЧЕСТВО', 'Комментарий', '', 'Устройство', '',
  'Время скана', 'Время записи', 'ID скана (не трогать)'];
var NUM_COLS = 11;
var COL_CODE = 1;   // A
var COL_DATE = 2;   // B
var COL_NAME = 3;   // C
var COL_SCAN_TIME = 9;   // I
var COL_WRITE_TIME = 10; // J
var COL_ID = 11;    // K
var VERSION = 4;

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (SECRET && params.token !== SECRET) {
    return json_({ ok: false, error: 'Неверный секретный ключ' });
  }
  var ss = getSpreadsheet_();
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

    var ss = getSpreadsheet_();
    var tz = ss.getSpreadsheetTimeZone();
    var rows = Array.isArray(data.rows) ? data.rows : [];
    var sheet = getSheet_(ss, data.sheet);
    var lastRow = lastRowInColumnA_(sheet);
    var known = existingIds_(sheet, lastRow);
    var now = serial_(new Date(), tz);

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
      var qty = Number(r.qty);
      if (!isFinite(qty) || qty === 0) qty = 1;
      var scanTime = serial_(ts, tz);
      values.push([
        String(r.code),              // A артикул
        Math.floor(scanTime),        // B дата без времени
        '',                          // C формула ставится ниже
        qty,                         // D количество
        String(r.note || ''),        // E комментарий
        '-',                         // F
        String(r.device || '-'),     // G устройство
        '-',                         // H
        scanTime,                    // I время скана
        now,                         // J время записи
        id                           // K ID скана
      ]);
      saved.push(id);
    });

    if (values.length) {
      var start = lastRow + 1;
      var n = values.length;
      if (start + n - 1 > sheet.getMaxRows()) {
        sheet.insertRowsAfter(sheet.getMaxRows(), start + n - 1 - sheet.getMaxRows());
      }
      // Артикул и ID как текст: иначе длинные коды превращаются в 4,6E+12 и теряются ведущие нули.
      sheet.getRange(start, COL_CODE, n, 1).setNumberFormat('@');
      sheet.getRange(start, COL_ID, n, 1).setNumberFormat('@');
      sheet.getRange(start, COL_DATE, n, 1).setNumberFormat('yyyy-mm-dd');
      sheet.getRange(start, COL_SCAN_TIME, n, 2).setNumberFormat('yyyy-mm-dd h:mm:ss');
      sheet.getRange(start, 1, n, NUM_COLS).setValues(values);
      fillNameFormulas_(ss, sheet, lastRow, start, n);
      SpreadsheetApp.flush();
    }

    return json_({ ok: true, saved: saved, duplicates: duplicates, timeZone: tz });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getSheet_(ss, name) {
  var sheetName = String(name || '').trim() || DEFAULT_SHEET;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (!sheet.getRange(1, COL_ID).getValue()) {
    sheet.getRange(1, COL_ID).setValue(HEADERS[COL_ID - 1]);
  }
  return sheet;
}

// Наименование в колонке C: та же формула, что и в строках выше, =ВПР(A…;'номенклатура'!A:B;2;ЛОЖЬ).
// Формулу копируем из последней рабочей строки: так она не зависит от языка таблицы
// (в русской таблице аргументы разделяются «;», в английской — «,»).
function fillNameFormulas_(ss, sheet, lastRow, start, n) {
  var target = sheet.getRange(start, COL_NAME, n, 1);
  if (lastRow >= 2) {
    var from = Math.max(2, lastRow - 300);
    var cells = sheet.getRange(from, COL_NAME, lastRow - from + 1, 1);
    var formulas = cells.getFormulas();
    var shown = cells.getDisplayValues();
    for (var i = formulas.length - 1; i >= 0; i--) {
      var f = formulas[i][0];
      if (f && /VLOOKUP|ВПР/i.test(f) && shown[i][0] !== '#ERROR!') {
        sheet.getRange(from + i, COL_NAME).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        return;
      }
    }
  }
  var locale = String(ss.getSpreadsheetLocale() || '');
  var sep = /^(en|ja|zh|ko|th|he)/.test(locale) ? ',' : ';';
  var list = [];
  for (var r = 0; r < n; r++) {
    list.push(['=VLOOKUP(A' + (start + r) + sep + "'" + NOMENCLATURE_SHEET + "'!A:B" + sep + '2' + sep + 'FALSE)']);
  }
  target.setFormulas(list);
}

// Последняя заполненная строка по колонке A (формулы или заметки в других колонках не сбивают).
function lastRowInColumnA_(sheet) {
  var max = sheet.getLastRow();
  if (max < 1) return 0;
  var col = sheet.getRange(1, COL_CODE, max, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (col[i][0] !== '' && col[i][0] !== null) return i + 1;
  }
  return 0;
}

function existingIds_(sheet, lastRow) {
  var map = {};
  if (lastRow < 2) return map;
  var ids = sheet.getRange(2, COL_ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0]) map[String(ids[i][0])] = true;
  }
  return map;
}

// Дата-время → число дней таблицы (как хранит Google Таблица) в часовом поясе таблицы.
function serial_(date, tz) {
  var p = Utilities.formatDate(date, tz, 'yyyy,MM,dd,HH,mm,ss').split(',').map(Number);
  var ms = Date.UTC(p[0], p[1] - 1, p[2], p[3], p[4], p[5]);
  return (ms - Date.UTC(1899, 11, 30)) / 86400000;
}

function getSpreadsheet_() {
  var id = String(SPREADSHEET_ID || '').trim();
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
