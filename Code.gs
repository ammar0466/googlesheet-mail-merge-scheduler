const CLIENT_ID = 'YOUR_CLIENT_ID';
const API_KEY = 'YOUR_API_KEY';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Start Mail Merge', 'showMailMergeDialog')
    .addToUi();
}

function showMailMergeDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('MailMerge')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mail Merge');
}

function getGmailDrafts() {
  const drafts = GmailApp.getDrafts();
  const draftData = drafts.map(draft => {
    return {
      id: draft.getId(),
      subject: draft.getMessage().getSubject()
    };
  });
  return draftData;
}


function getColumnHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}



function sendEmails(columnIndex, draftId, schedule, scheduledDateTime) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const headers = getColumnHeaders();
  const mergeStatusHeader = 'Merge Status';
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  const mergeStatusIndex = headers.indexOf(mergeStatusHeader);

  if (mergeStatusIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(mergeStatusHeader);
    dataRange.setValues(data);
  }

  const mergeStatusColumn = sheet.getRange(2, mergeStatusIndex + 1, lastRow - 1).getValues();
  const draft = GmailApp.getDraft(draftId);
  const message = draft.getMessage();
  const messageBody = message.getBody();

  for (let i = 0; i < data.length; i++) {
    const currentMergeStatus = mergeStatusColumn[i][0].toString();
    if (currentMergeStatus === 'sent' || currentMergeStatus === '0' || currentMergeStatus.includes('Scheduled on')) {
      continue;
    }

    const email = data[i][columnIndex];
    let customizedBody = messageBody;
    let customizedSubject = message.getSubject();
    headers.forEach((header, index) => {
      const value = data[i][index];
      customizedBody = replaceAll(customizedBody, `{{${header}}}`, value);
      customizedSubject = replaceAll(customizedSubject, `{{${header}}}`, value);
    });
    const mailOptions = {
      subject: customizedSubject,
      htmlBody: customizedBody
    };

    if (schedule) {
      const scheduledEmailData = {
        rowIndex: i,
        email: email,
        columnIndex: columnIndex,
        draftId: draftId,
        scheduledDateTime: scheduledDateTime
      };
      const trigger = ScriptApp.newTrigger('sendScheduledEmail')
        .timeBased()
        .at(new Date(scheduledDateTime))
        .create();

      // Store the trigger ID and scheduled email data in the Google Sheet
      sheet.getRange(i + 2, mergeStatusIndex + 2).setValue(trigger.getUniqueId());
      sheet.getRange(i + 2, mergeStatusIndex + 3).setValue(JSON.stringify(scheduledEmailData));

      sheet.getRange(i + 2, mergeStatusIndex + 1).setValue(`Scheduled on ${scheduledDateTime}`);
    }  else {
      GmailApp.sendEmail(email, mailOptions.subject, "", mailOptions);
      sheet.getRange(i + 2, mergeStatusIndex + 1).setValue('sent');
    }
  }
}




// function replaceAll(str, find, replace) {
//   return str.split(find).join(replace);
// }
function replaceAll(str, find, replace) {
  const escapedFind = find.replace(/[.*+\-?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(escapedFind, 'g');
  return str.replace(regex, replace);
}


function sendScheduledEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const mergeStatusHeader = 'Merge Status';
  const mergeStatusIndex = getColumnHeaders().indexOf(mergeStatusHeader);
  const lastRow = sheet.getLastRow();
  const triggerIdColumn = sheet.getRange(2, mergeStatusIndex + 2, lastRow - 1).getValues();
  const emailDataColumn = sheet.getRange(2, mergeStatusIndex + 3, lastRow - 1).getValues();

  const now = new Date();

  for (let i = 0; i < triggerIdColumn.length; i++) {
    const emailData = emailDataColumn[i][0];
    if (!emailData) {
      continue;
    }

    const scheduledEmailData = JSON.parse(emailData);
    const scheduledDateTime = new Date(scheduledEmailData.scheduledDateTime);

    // Skip if the difference is more than 1 minute
    if (Math.abs(scheduledDateTime - now) > 60000) {
      continue;
    }

    const rowIndex = scheduledEmailData.rowIndex;
    const email = scheduledEmailData.email;
    const draftId = scheduledEmailData.draftId;

    const draft = GmailApp.getDraft(draftId);
    const message = draft.getMessage();
    const messageBody = message.getBody();
    let customizedBody = messageBody;

    getColumnHeaders().forEach((header, index) => {
      const value = sheet.getRange(rowIndex + 2, index + 1).getValue();
      customizedBody = replaceAll(customizedBody, `{{${header}}}`, value);
    });

    const mailOptions = {
      subject: message.getSubject(),
      htmlBody: customizedBody
    };

    GmailApp.sendEmail(email, mailOptions.subject, "", mailOptions);
    sheet.getRange(rowIndex + 2, mergeStatusIndex + 1).setValue('sent');

    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === scheduledEmailData.triggerId) {
        Logger.log('Deleting trigger: ' + trigger.getUniqueId());
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
}








function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('MailMerge');
}

function getTriggerBalance() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.length;
}

function deleteOldTriggers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const triggers = ScriptApp.getProjectTriggers();
  const now = new Date();
  const headers = getColumnHeaders();
  const mergeStatusIndex = headers.indexOf('Merge Status');
  const triggerIdColumn = sheet.getRange(2, mergeStatusIndex + 2, sheet.getLastRow() - 1).getValues();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'sendScheduledEmail') {
      const triggerUniqueId = trigger.getUniqueId();

      for (let i = 0; i < triggerIdColumn.length; i++) {
        const currentTriggerId = triggerIdColumn[i][0];

        if (currentTriggerId === triggerUniqueId) {
          const mergeStatus = sheet.getRange(i + 2, mergeStatusIndex + 1).getValue();

          if (mergeStatus === 'sent') {
            ScriptApp.deleteTrigger(trigger);
          }

          break;
        }
      }
    }
  });
}



