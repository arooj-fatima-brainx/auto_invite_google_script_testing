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
  var columnNames = getColumnNames(); // Get the column names
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var template = HtmlService.createTemplateFromFile('Dialog');
  template.columnNames = columnNames;
  template.spreadsheetId = spreadsheetId; // Pass the spreadsheet ID to the template
  var html = template.evaluate().setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, DIALOG_TITLE);
}

function getColumnNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const columnNames = values[0]; // Assumes the first row contains column names
  return columnNames;
}

function getColumnData(columnName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var columnIndex;
  if (headers.indexOf(columnName) !== -1) {
    columnIndex = headers.indexOf(columnName);
  } else {
    columnIndex = 0;
  }
  var dataRange = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1);
  var columnData = dataRange.getValues().flat();
  return columnData;
}

function sendInvitation(calendarId, eventId, attendeeEmails) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error('Calendar not found');
    }

    var eveId =  decodeBase64(eventId).split(' ')[0];
    var event = Calendar.Events.get(calendarId, eveId);
    if (!event) {
      throw new Error('event not found');
    }

    if (!event.attendees) {
      event.attendees = [];
    }

    attendeeEmails.forEach(function(email) {
      var isEmailPresent = event.attendees.some(function(attendee) {
        return attendee.email === email;
      });
      if (!isEmailPresent) {
        event.attendees.push({ email: email });
      }
    });
    event = Calendar.Events.patch(event, calendarId, eveId, {
      sendNotifications: true
    });

    return "SUCCESS";
  } catch (error) {
    throw new Error(error.message);
  }
}

function decodeBase64(encodedString) {
   var decodedBytes = Utilities.base64Decode(encodedString);
  var decodedString = Utilities.newBlob(decodedBytes).getDataAsString("UTF-8");
  return decodedString;
}