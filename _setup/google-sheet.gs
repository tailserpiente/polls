// Приёмник ответов анкет для Google-таблицы.
// Каждая анкета — новая строка; недостающие столбцы добавляются автоматически по названиям вопросов.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const lastCol = sheet.getLastColumn();
    const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) headers.push(k);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.appendRow(headers.map(function (h) { return data[h] !== undefined ? String(data[h]) : ""; }));
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
