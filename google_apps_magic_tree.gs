/**
 * «Волшебное дерево» — приём ответов детей в Google Таблицу.
 *
 * 1. Вставьте ID своей таблицы в SHEET_ID.
 * 2. Разверните скрипт как Web app:
 *    Execute as: Me
 *    Who has access: Anyone
 * 3. Скопируйте полученный Web app URL в HTML-файл.
 */

const SHEET_ID = 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ_СЮДА';
const SHEET_NAME = 'Ответы';

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Получен пустой запрос');
    }

    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    prepareSheet_(sheet);

    sheet.appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.tree),
      clean_(data.prediction),
      clean_(data.finalFruit),
      clean_(data.observation)
    ]);

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
    'Имя ребёнка',
    'Номер дерева',
    'Прогноз',
    'Что осталось в конце',
    'Догадка ребёнка'
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
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 190);
  sheet.setColumnWidth(6, 420);
  sheet.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm:ss');
  sheet.getRange('F:F').setWrap(true);
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
