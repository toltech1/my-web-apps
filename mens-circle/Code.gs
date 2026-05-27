// =============================================
// Men's Circle Feedback App — Google Apps Script API
// Paste this into your Apps Script project (replaces the old Code.gs)
// =============================================

var SHEET_NAME = 'Feedback';
var MEMBERS = ['Ben', 'David', 'Guillaume', 'Pablo', 'Jules', 'Hans', 'Thomas'];

// Handle all GET requests (used as REST API by the frontend)
function doGet(e) {
  var action = e.parameter.action;
  var result;

  try {
    if (action === 'submit') {
      result = submitFeedback({
        from: e.parameter.from || '',
        to:   e.parameter.to   || '',
        appreciation: e.parameter.q1 || '',
        growth:       e.parameter.q2 || '',
        word:         e.parameter.q3 || '',
        challenge:    e.parameter.q4 || ''
      });
    } else if (action === 'getSubmitted') {
      result = getSubmittedByMember(e.parameter.from);
    } else if (action === 'getFeedback') {
      result = getFeedbackForMember(e.parameter['for']);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Save feedback to the sheet
function submitFeedback(data) {
  var sheet = getOrCreateSheet();
  var month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  sheet.appendRow([
    new Date(),
    data.from,
    data.to,
    data.appreciation,
    data.growth,
    data.word,
    data.challenge,
    month
  ]);
  return { success: true };
}

// Get list of names this person has already given feedback to this month
function getSubmittedByMember(fromName) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var submitted = [];
  var month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === fromName && data[i][7] === month) {
      submitted.push(data[i][2]);
    }
  }
  return submitted;
}

// Get all feedback received by a member this month
function getFeedbackForMember(name) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var results = [];
  var month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === name && data[i][7] === month) {
      results.push({
        from:         data[i][1],
        appreciation: data[i][3],
        growth:       data[i][4],
        word:         data[i][5],
        challenge:    data[i][6]
      });
    }
  }
  return results;
}

// Helper: get or create the Feedback sheet
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp','From','To','Appreciation','Growth Edge','One Word','Challenge','Month']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}
