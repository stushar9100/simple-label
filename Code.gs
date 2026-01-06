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
  var columnsPerRow = Math.max(1, Number(options && options.columnsPerRow) || 1);
  var labelPadding = Math.max(0, Number(options && options.labelPadding) || 8);
  var labelSpacing = Math.max(0, Number(options && options.labelSpacing) || 0);

  var doc = DocumentApp.create('Labels - ' + sheet.getName());
  var body = doc.getBody();
  body.clear();

  var table = body.appendTable();
  if (table.getNumRows() > 0) {
    table.removeRow(0);
  }
  table.setAttributes({
    [DocumentApp.Attribute.BORDER_COLOR]: '#bdbdbd',
    [DocumentApp.Attribute.BORDER_WIDTH]: 1,
    [DocumentApp.Attribute.BORDER_STYLE]: 'DOTTED'
  });

  function styleLabelCell(cell) {
    cell.setPaddingTop(labelPadding);
    cell.setPaddingBottom(labelPadding);
    cell.setPaddingLeft(labelPadding);
    cell.setPaddingRight(labelPadding);
  }

  function styleSpacerCell(cell) {
    cell.setPaddingTop(0);
    cell.setPaddingBottom(0);
    cell.setPaddingLeft(0);
    cell.setPaddingRight(0);
    cell.clear();
    if (labelSpacing > 0) {
      cell.setWidth(labelSpacing);
    }
  }

  var totalColumns = columnsPerRow + (labelSpacing > 0 ? columnsPerRow - 1 : 0);
  var currentRow = null;

  values.slice(1).forEach(function (row, rowIndex) {
    var rowValues = orderedHeaders.map(function (header) {
      var index = headerIndex[header];
      return index === undefined ? '' : row[index];
    });
    var positionInRow = rowIndex % columnsPerRow;
    if (positionInRow === 0) {
      if (labelSpacing > 0 && rowIndex > 0) {
        var spacerRow = table.appendTableRow();
        for (var spacerIndex = 0; spacerIndex < totalColumns; spacerIndex += 1) {
          styleSpacerCell(spacerRow.appendTableCell());
        }
        spacerRow.setMinimumHeight(labelSpacing);
      }
      currentRow = table.appendTableRow();
      for (var cellIndex = 0; cellIndex < totalColumns; cellIndex += 1) {
        currentRow.appendTableCell('');
      }
      currentRow.setMinimumHeight(72);
    }

    var cellPosition = positionInRow * (labelSpacing > 0 ? 2 : 1);
    var cell = currentRow.getCell(cellPosition);
    cell.clear();
    styleLabelCell(cell);

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

    if (labelSpacing > 0 && cellPosition + 1 < totalColumns) {
      var spacerCell = currentRow.getCell(cellPosition + 1);
      styleSpacerCell(spacerCell);
    }
  });

  doc.saveAndClose();
  return doc.getUrl();
}
