var ROW_OFFSET = 2;

function addToCombined(sheet, index) {
  var generated = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2WK-generated"); // Put the tab name here

  var numRows = sheet.getLastRow();
  var filledRows = 0;
  var skippedRows = 0;
  for (var i = 0; i < numRows - ROW_OFFSET; i++) {
    var whatContent = sheet.getRange(i+1+ROW_OFFSET, 3);
    var timeContent = sheet.getRange(i+1+ROW_OFFSET, 1);
    if (!whatContent.isBlank() || !timeContent.isBlank()) {
      var range = generated.getRange(i + index + 1 + ROW_OFFSET - skippedRows, 1, 1, sheet.getLastColumn() - 2);
      if (whatContent.isBlank()) {
        range.setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build());
      }
      
      var data = sheet.getRange(i+1 + ROW_OFFSET, 1, 1, sheet.getLastColumn() - 2);
      var formulaRange = sheet.getRange(i+1+ROW_OFFSET, 3);
      var linkUrl = formulaRange.getRichTextValue().getLinkUrl()

      range.setValues(data.getValues());
      if (linkUrl) {
        var formulaToRange = generated.getRange(i+index+1+ROW_OFFSET-skippedRows, 3);
        var textValue = formulaToRange.getValue();
        var newRichText = SpreadsheetApp.newRichTextValue().setText(formulaToRange.getValue()).setLinkUrl(linkUrl).build();
        formulaToRange.setRichTextValue(newRichText);
      }
      filledRows++;
    } else {
      skippedRows++;
    }
  }

  return filledRows;
}

function combineTabs() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var generated = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2WK-generated"); // Put the tab name here
  var range = generated.getRange(3, 1, generated.getLastRow(), generated.getLastColumn());
  range.clear();

  var timeRange = generated.getRangeList(['A1:A', 'B1:B']);
  timeRange.setNumberFormat('h:mm am/pm')

  var allGenerated = generated.getRangeList(['A1:H']);
  allGenerated.setFontFamily('Arial Narrow');
  if (sheets.length == 0) {
    return;
  }

  var row = 0;
  for (var index = 0; index < sheets.length; index++) {
    var sheet = sheets[index];
    var name = sheet.getName();
    if (containsDate(name)) {
      Utilities.sleep(1000);
      var curRange = generated.getRange(row + 1 + ROW_OFFSET, 1);
      curRange.setValue(name);
      curRange.setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build());
      curRange = generated.getRange(row+ 1 + ROW_OFFSET, 1, 1, generated.getLastColumn());
      curRange.setBackground('#d9ead3');
      row += 1;

      var filledRows = addToCombined(sheet, row);
      row += filledRows;
      row += 1;
    }
  }
}

function containsDate(name) {
  var re = /^\w{3} \d+\/\d+$/;
  return re.test(name);
}
