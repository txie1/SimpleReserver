function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Reservation')
    .addItem('Open Reservation Panel', 'openSidebar')
    .addItem('Clear Yesterday', 'clearDailyReservations') // manual trigger
    .addItem('Clear All Entries', 'clearAllEntries')      // manual trigger
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ReservationPanel')
      .setTitle('Resource Reservation Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getActiveRange() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  return range ? range.getA1Notation() : null;
}

function confirmUsage(color, name, resourceCount, comment) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert("No cells selected. Please select cells to confirm usage.");
    return;
  }

  // Fetch day & time reserved 
  var column = range.getColumn();
  var day = sheet.getRange(1, column).getValue();     // Get column header
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var timeRange = sheet.getRange(startRow, 1, numRows, 1).getValues();
  var startTime = timeRange[0][0];
  var endTime = timeRange[timeRange.length - 1][0];
  var formattedRange = formatTimeRange(startTime, endTime);

  // Confirmation
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    `[Confirming Reservation]`,
    `Name: ${name}\nCount: ${resourceCount}\nTime: ${formattedRange}\nPurpose/Notes: ${comment}\n\nClick OK to notify users.`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response != ui.Button.OK) {
    ui.alert("Reservation cancelled.");
    return;
  }

  // Update cells with [name, resource count, color, comment]
  var nameWithResource = resourceCount ? `${name} [${resourceCount}]` : name;
  range.setValue(nameWithResource);
  range.setBackground(color);
  if (comment) {
    range.setComment(comment);
  }
  sendEmailNotification(name, resourceCount, day, formattedRange, comment);
}


function sendEmailNotification(name, resourceCount, day, formattedRange, comment) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(spreadsheet.getId());
  var editors = file.getEditors();
  var recipientEmails = editors.map(user => user.getEmail());
  if (recipientEmails.length === 0) return;

  // Email content
  var subject = "New Resource Reservation";
  var body = `
    A new reservation has been made:<br>
    <br>
    <strong>User:</strong> ${name}<br>
    <strong>Count:</strong> ${resourceCount}<br>
    <strong>Date:</strong> ${day}<br>
    <strong>Time:</strong> ${formattedRange}<br>
    <strong>Purpose/Notes:</strong> ${comment}<br>
    <br>
    Please check the <a href="${spreadsheet.getUrl()}">Reservation Sheet</a> for more details.
  `;

  GmailApp.sendEmail(recipientEmails.join(","), subject, "", { htmlBody: body });
}

function formatTimeRange(startTime, endTime) {
  if (!startTime || !endTime) return "Time range unavailable";
  return `${Utilities.formatDate(startTime, Session.getScriptTimeZone(), "h:mm a")} - ${Utilities.formatDate(endTime, Session.getScriptTimeZone(), "h:mm a")}`;
}



function clearDailyReservations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 2, 1, 7).getValues()[0];
  var yesterday = getYesterdayDay();
  var columnIndex = headers.indexOf(yesterday);
  
  if (columnIndex === -1) {
    Logger.log("Day not found in headers. Nothing cleared.");
    return;
  }

  var columnToClear = columnIndex + 2;
  var range = sheet.getRange(2, columnToClear, 19, 1);
  range.clearContent();
  range.clearNote();
  range.setBackground("#b7b7b7");
  Logger.log(`Cleared reservations for ${yesterday}.`);
}

function getYesterdayDay() {
  var today = new Date();
  today.setDate(today.getDate() - 1);
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  return days[today.getDay()];
}

function resetWeeklyReservations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange('B2:H20');
  range.clearContent();
  range.clearNote();
  range.setBackground(null);
  Logger.log("Reservations sheet has been successfully reset.");
}



function clearAllEntries() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Clear All Reservations',
    'Are you sure you want to clear ALL reservations in the table?',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange('B2:H20');
    range.clearContent();
    range.clearNote();
    range.setBackground(null);
  } else {
    ui.alert('Action cancelled.');
  }
}
