const APP_TITLE = 'Auto Invite Setup';
var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var userProperties = PropertiesService.getUserProperties();
var activeSheet = spreadSheet.getActiveSheet();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Auto Invite App')
    .addItem('Set Invites Configurations', 'showInviteSetupModal')
    .addToUi();
}

function showInviteSetupModal() {
  var html = HtmlService.createTemplateFromFile('SetupDialog').evaluate().setWidth(600).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, APP_TITLE);
}

function checkAutoInviteSetup() {
  var value = userProperties.getProperty("eventId");
  return value != null;
}

function getColumnNamesWithReferences() {
  var firstRowValues = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
  let columnNamesWithReferences = [];
  for (let i = 0; i < firstRowValues.length; i++) {
    columnNamesWithReferences.push({ name: firstRowValues[i], reference: String.fromCharCode(65 + i) });
  }
  return columnNamesWithReferences;
}

function getEmailColumnData(columnReference) {
 try {
    let range = activeSheet.getRange(`${columnReference}2:${columnReference}${activeSheet.getLastRow()}`);
    let emails = range.getValues().flat();
    emails = emails.filter((email) => {
      if(isEmail(email)){
        return email;
      }
    });
    return emails;
  } catch (error) {
    throw new Error(error.message);
  }
}

function sendInvitation(eventId, attendeeEmails) {
  try {
    let calendarId = getCurrentUserEmail();
    let eveId =  decodeBase64(eventId).split(' ')[0];
    var event = Calendar.Events.get(calendarId, eveId);

    if (!event) {
      throw new Error(`Since the Calendar Event you're linking to is owned by someone else, they'll need to first add your email ${calendarId} as a guest on the Calendar Event. Once you're listed as a guest on the Calendar, this app will be able to successfully send out invites.`);
    }

    if (!event.attendees) {
      event.attendees = [];
    }

    attendeeEmails.forEach((email) => {
      if(!isEmail(email) || event.attendees.includes(email)){
        return
      }
      event.attendees.push({ email: email });
    });

    event = Calendar.Events.patch(event, calendarId, eveId, {
      sendNotifications: true
    });

    return "SUCCESS";
  } catch (error) {
    throw new Error(error.message);
  }
}

function handleFormSubmitEvent(e){
  let refCol = userProperties.getProperty("emailReferenceColumn")
  let eventId = userProperties.getProperty("eventId")

  if ( refCol && eventId){
    let emailsArray = []
    emailsArray.push(activeSheet.getRange(`${refCol}${e.range.getRow()}`).getValue());
    sendInvitation(eventId, emailsArray);
  }
}

// helper methods

function createFormSubmitTrigger() {
  try {
    let trigger = ScriptApp.newTrigger('handleFormSubmitEvent')
      .forSpreadsheet(spreadSheet)
      .onFormSubmit()
      .create();
    userProperties.setProperty("triggerId", trigger.getUniqueId());
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteFormSubmitTrigger() {
    try {
      let allTriggers = ScriptApp.getProjectTriggers();
      let triggerId = userProperties.getProperty("triggerId");
        for (let i = 0; i < allTriggers.length; i++) {
          let trigger = allTriggers[i];
          if (trigger.getUniqueId() === triggerId) {
            ScriptApp.deleteTrigger(trigger);
            return;
          }
        }
      } catch(error) {
      throw new Error(error.message);
    }
}

function storeEventInfo(eventId, emailReferenceColumn){
  userProperties.setProperty("eventId", eventId);
  userProperties.setProperty("emailReferenceColumn", emailReferenceColumn);
}

function removeEventInfo() {
  userProperties.deleteProperty("eventId");
  userProperties.deleteProperty("emailReferenceColumn");
  userProperties.deleteProperty("triggerId");
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();;
}

function decodeBase64(encodedString) {
  let decodedBytes = Utilities.base64Decode(encodedString);
  let decodedString = Utilities.newBlob(decodedBytes).getDataAsString("UTF-8");
  return decodedString;
}

function isEmail(email) {
  let regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}
