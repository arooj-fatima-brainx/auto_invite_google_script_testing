var DIALOG_TITLE = 'Auto Invite Setup';

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
    'Let\'s get to know each other!',
    'Please enter your name:',
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Your name is ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Menu')
    .addItem('Show dialog', 'showDialog')
    .addToUi();
}

function showDialog() {
  console.log('in show dialog');
  var columnNames = getColumnNames(); // Get the column names
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var template = HtmlService.createTemplateFromFile('Dialog');
  template.columnNames = columnNames;
  template.spreadsheetId = spreadsheetId; // Pass the spreadsheet ID to the template
  var html = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, DIALOG_TITLE);
}

function getColumnNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const columnNames = values[0]; // Assumes the first row contains column names
  return columnNames;
}

function getColumnIndexByName(spreadsheetId, sheetName, columnName) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == columnName) {
      return i + 1; // Adding 1 because getColumn() is 1-indexed
    }
  }
  // If column name not found, return -1
  return -1;
}
