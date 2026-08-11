/**
 * «Волшебное дерево» — приём ответов детей в Google Таблицу.
 *
 * 1. Вставьте ID своей таблицы в SHEET_ID.
 * 2. Разверните скрипт как Web app:
 *    Execute as: Me
 *    Who has access: Anyone
 * 3. Скопируйте полученный Web app URL в HTML-файл.
 */

const SHEET_ID = '14g0-TXfVLjbk1YG7TAtHjWKG_v7YVmZV9k70CsUAjpk';
const SHEET_NAME = 'Ответы';
const GUESSES_SHEET_NAME = 'Догадки';

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Получен пустой запрос');
    }

    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const isSavedAnswer = clean_(data.event) === 'Ответ сохранён';
    let sheet = spreadsheet.getSheetByName(isSavedAnswer ? GUESSES_SHEET_NAME : SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(isSavedAnswer ? GUESSES_SHEET_NAME : SHEET_NAME);
    }

    if (isSavedAnswer) {
      prepareGuessesSheet_(sheet);
      sheet.appendRow([
        new Date(),
        clean_(data.name),
        clean_(data.tree),
        clean_(data.startBanana),
        clean_(data.startPineapple),
        clean_(data.prediction),
        clean_(data.attempt),
        clean_(data.finalFruit),
        clean_(data.observation),
        clean_(data.sessionId),
        clean_(data.page)
      ]);
    } else {
      prepareSheet_(sheet);
      sheet.appendRow([
        new Date(),
        clean_(data.name),
        clean_(data.event),
        clean_(data.tree),
        clean_(data.startBanana),
        clean_(data.startPineapple),
        clean_(data.prediction),
        clean_(data.attempt),
        clean_(data.action),
        clean_(data.state),
        clean_(data.finalFruit),
        clean_(data.observation),
        clean_(data.sessionId),
        clean_(data.page),
        clean_(data.source || 'volshebnoe-derevo')
      ]);
    }

    return jsonResponse_({ ok: true });
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return jsonResponse_({
    ok: true,
    message: 'Скрипт «Волшебное дерево» работает'
  });
}

function prepareSheet_(sheet) {
  if (sheet.getLastRow() > 0) {
    return;
  }

  const headers = [[
    'Дата и время',
    'Имя и фамилия',
    'Событие',
    'Номер дерева',
    'Старт: бананы',
    'Старт: ананасы',
    'Прогноз',
    'Попытка / ход',
    'Действие ученика',
    'Текущее состояние',
    'Что осталось в конце',
    'Догадка ребёнка',
    'ID сессии',
    'Страница',
    'Источник'
  ]];

  const headerRange = sheet.getRange(1, 1, 1, headers[0].length);
  headerRange
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#52327c')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 240);
  sheet.setColumnWidth(10, 240);
  sheet.setColumnWidth(12, 420);
  sheet.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm:ss');
  sheet.getRange('I:L').setWrap(true);
}

function prepareGuessesSheet_(sheet) {
  if (sheet.getLastRow() > 0) {
    return;
  }

  const headers = [[
    'Дата и время',
    'Имя и фамилия',
    'Номер дерева',
    'Старт: бананы',
    'Старт: ананасы',
    'Прогноз',
    'Попытка / ход',
    'Что осталось в конце',
    'Догадка ребёнка',
    'ID сессии',
    'Страница'
  ]];

  const headerRange = sheet.getRange(1, 1, 1, headers[0].length);
  headerRange
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#52327c')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(9, 420);
  sheet.setColumnWidth(10, 220);
  sheet.setColumnWidth(11, 220);
  sheet.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm:ss');
  sheet.getRange('I:I').setWrap(true);
}

function clean_(value) {
  if (value === undefined || value === null) {
    return '';
  }

  const text = String(value).trim();

  // Не даём значениям из формы превращаться в формулы Google Таблиц.
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
