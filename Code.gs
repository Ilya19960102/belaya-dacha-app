/* Приём результатов тестов и запросов кода из приложения «Инструкция Белая Дача».
   Вставь этот файл в Google Apps Script, привязанный к своей Google Таблице
   (Таблица → Расширения → Apps Script), затем: Начать развертывание → Новое развертывание →
   тип «Веб-приложение», «Выполнять от имени: меня», «Доступ: все». Адрес /exec вставь в config.js. */

// ДОЛЖНА совпадать с SALT в unlock.js
var CODE_SALT = 'bd-tests-2026-8DH8nwjN';

var SHEET_RESULTS = 'Результаты';
var SHEET_REQUESTS = 'Запросы кода';

function hash32_(s) {
  var h = 0x811c9dc5;
  for (var i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i) & 0xff;
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

// Номер в любом виде → 7XXXXXXXXXX (или '' если не похоже на номер)
function digits_(phone) {
  var d = String(phone || '').replace(/\D/g, '');
  if (d.length === 11 && (d[0] === '7' || d[0] === '8')) d = d.slice(1);
  return d.length === 10 ? '7' + d : '';
}

function codeFor_(phone) {
  var d = digits_(phone);
  if (!d) return '';
  return ('0000' + (hash32_(CODE_SALT + ':' + d) % 10000)).slice(-4);
}

/** Код доступа к тестам по номеру телефона. В таблице: =КОД(B2)
 * @customfunction */
function КОД(phone) { return codeFor_(phone); }

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(header); sh.setFrozenRows(1); sh.getRange(1, 1, 1, header.length).setFontWeight('bold'); }
  return sh;
}

// Телефон с повторной отправкой может прислать ту же запись дважды — пропускаем дубли по rid
function seen_(sh, rid, col) {
  if (!rid || sh.getLastRow() < 2) return false;
  return sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(String(rid)).matchEntireCell(true).findNext() !== null;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var r = JSON.parse(e.postData.contents);
    var when = new Date(Number(r.ts) || Date.now());
    if (r.action === 'coderequest') {
      var sq = sheet_(SHEET_REQUESTS, ['Время', 'ФИО', 'Телефон', 'Код для сотрудника', 'Устройство', 'rid']);
      if (!seen_(sq, r.rid, 6)) sq.appendRow([when, r.fio || '', r.tab || '', "'" + codeFor_(r.tab), r.device || '', r.rid || '']);
    } else {
      var sr = sheet_(SHEET_RESULTS, ['Время', 'ФИО', 'Телефон', 'Тест', 'Верно', 'Всего', '%', 'Экзамен сдан', 'Длительность, с', 'Ошибки', 'Устройство', 'Версия', 'rid']);
      if (!seen_(sr, r.rid, 13)) sr.appendRow([when, r.fio || '', r.tab || '', r.type || '', r.correct, r.total, r.percent,
        r.passed === null || r.passed === undefined ? '' : (r.passed ? 'да' : 'нет'),
        r.duration || '', (r.mistakes || []).join('\n'), r.device || '', r.app || '', r.rid || '']);
    }
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ContentService.createTextOutput('Белая Дача: приём результатов работает'); }
