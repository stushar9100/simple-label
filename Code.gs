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

function createLabelsDoc(options) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getDisplayValues();
  if (values.length < 2) {
    throw new Error('No data rows found.');
  }

  var headers = values[0];
  var headerIndex = {};
  headers.forEach(function (header, index) {
    if (header && header.toString().trim() !== '') {
      headerIndex[header.toString()] = index;
    }
  });

  var orderedHeaders = options && options.order ? options.order : [];
  if (orderedHeaders.length === 0) {
    throw new Error('Select at least one column.');
  }

  var textSize = Number(options && options.textSize) || 12;
  var lineSpacing = Number(options && options.lineSpacing) || 4;
  var lineBreaks = options && options.lineBreaks !== false;

  var doc = DocumentApp.create('Labels - ' + sheet.getName());
  var body = doc.getBody();
  body.clear();

  var table = body.appendTable();
  if (table.getNumRows() > 0) {
    table.removeRow(0);
  }
  values.slice(1).forEach(function (row) {
    var rowValues = orderedHeaders.map(function (header) {
      var index = headerIndex[header];
      return index === undefined ? '' : row[index];
    });
    var labelRow = table.appendTableRow();
    labelRow.setMinimumHeight(72);
    var cell = labelRow.appendTableCell();
    cell.clear();
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(8);
    cell.setPaddingRight(8);

    if (lineBreaks) {
      rowValues.forEach(function (value, index) {
        var paragraph = cell.appendParagraph(value || '');
        paragraph.setFontFamily('Arial');
        paragraph.setFontSize(textSize);
        paragraph.setSpacingAfter(index === rowValues.length - 1 ? 0 : lineSpacing);
      });
    } else {
      var paragraph = cell.appendParagraph(rowValues.join(' • '));
      paragraph.setFontFamily('Arial');
      paragraph.setFontSize(textSize);
      paragraph.setSpacingAfter(0);
    }
  });

  doc.saveAndClose();
  return doc.getUrl();
}
