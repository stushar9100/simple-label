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
