function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Label Generator')
    .addItem('Open', 'showLabelDialog')
    .addToUi();
}

function showLabelDialog() {
  var html = HtmlService.createHtmlOutputFromFile('labelDialog')
    .setWidth(600)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Create Labels');
}

function getSheetHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return [];
  }
  var values = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues();
  return values[0].filter(function (header) {
    return header && header.toString().trim() !== '';
  });
}
