/**
 * Habibi — приём заявок с лендинга в Google Таблицу.
 *
 * Установка см. в README.md.
 */

// Имя листа, куда будут писаться заявки
const SHEET_NAME = 'Leads';

function doPost(e) {
  try {
    const sheet = getSheet_();

    // Данные приходят как JSON-строка в e.postData.contents
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),                 // дата и время заявки
      data.name || '',
      data.contact || '',
      data.interest || '',
      data.lang || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Возвращает лист "Leads", создавая его и заголовки при первом запуске
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Date', 'Name', 'Contact', 'Interest', 'Language']);
    sheet.setFrozenRows(1);
  }

  return sheet;
}
