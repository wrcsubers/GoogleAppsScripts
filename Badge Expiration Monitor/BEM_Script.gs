// =========================================================================================================================
// ATLS Badge Expiration Monitor
// Runs at the beginning of each month to notify users of expiring access badges
// Written December 2023 by Cameron Woods - Cameron.Woods@Aristocrat.com
// =========================================================================================================================

// Add button to toolbar for manually running check function
// ====================================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Functions').addItem('Perform Badge Check', 'checkBadges').addToUi();
}

// Function for checking badge expiration dates and sending email
// ====================================================================================================
function checkBadges() {

  // Variables for function
  var spreadsheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var todaysDate = new Date();
  var todayPlus30 = new Date();
  var todayPlus60 = new Date();
  var todayPlus90 = new Date();
  todaysDate.setHours(0, 0, 0, 0);
  var todayPlus30 = todayPlus30.setDate(todaysDate.getDate() + 30);
  var todayPlus60 = todayPlus60.setDate(todaysDate.getDate() + 60);
  var todayPlus90 = todayPlus90.setDate(todaysDate.getDate() + 90);

  // Perform Checks on each sheet in spreadsheet
  for (var i = 0; i < spreadsheets.length; i++) {

    // Variables for each sheet
    var emailAddress = spreadsheets[i].getRange("B1").getValue();
    var firstName = emailAddress.toString().split('.')[0];
    var expiration00 = [];
    var expiration30 = [];
    var expiration60 = [];
    var expiration90 = [];
    var rowUnmonitoredCount = 0;
    var sheetErrors = false;

    // 
    if ((emailAddress == null) || (emailAddress == '')) {
      Logger.log("FATAL ERROR - Email Address Not Found on: " + spreadsheets[i].getSheetName());
      break;
    }

    // Perform Checks on each row in sheet
    for (var j = 5; j <= spreadsheets[i].getLastRow(); j++) {

      // Variables for each row
      var rowBadgeValue = spreadsheets[i].getRange("A" + j).getValue();
      var rowDateValue = spreadsheets[i].getRange("B" + j).getValue().toLocaleString().split(',')[0];
      var rowDateMillis = Date.parse(rowDateValue);
      var rowMonitor = spreadsheets[i].getRange("C" + j).getValue().toString();

      // Check to see rows are filled out correctly
      if ((rowBadgeValue != '') && (rowDateValue != '')) {
        // Only check row if it's enabled
        if (rowMonitor == '') {
          Logger.log("WARNING ERROR - Missing Row Monitor: " + spreadsheets[i].getSheetName() + ", row: " + j);
          sheetErrors = true;
        } else if (rowMonitor == 'true') {
          // Check Days from Expiration, if within range, add data to array
          if (rowDateMillis <= todaysDate) {                       // Past Expired
            expiration00.push([rowBadgeValue, rowDateValue]);
          } else if (rowDateMillis <= todayPlus30) {               // Within 30 Days of Expiration
            expiration30.push([rowBadgeValue, rowDateValue]);
          } else if (rowDateMillis <= todayPlus60) {               // Within 60 Days of Expiration
            expiration60.push([rowBadgeValue, rowDateValue]);
          } else if (rowDateMillis <= todayPlus90) {               // Within 90 Days of Expiration
            expiration90.push([rowBadgeValue, rowDateValue]);
          }
        } else if (rowMonitor == 'false') {
          rowUnmonitoredCount++;
        }
      } else {
        Logger.log("WARNING ERROR - Incorrect Row Format: " + spreadsheets[i].getSheetName() + ", row: " + j);
        sheetErrors = true;
      }
    }

    // Count number of expiring badges
    var numBadgesExpiring = expiration00.length + expiration30.length + expiration60.length + expiration90.length;

    // Load email template in
    var emailTemplate = HtmlService.createTemplateFromFile("EmailTemplate.html");

    // Assign variables in email template
    emailTemplate.numBadgesExpiring = numBadgesExpiring;
    emailTemplate.firstName = firstName;
    emailTemplate.expiration00 = expiration00;
    emailTemplate.expiration30 = expiration30;
    emailTemplate.expiration60 = expiration60;
    emailTemplate.expiration90 = expiration90;
    emailTemplate.rowUnmonitoredCount = rowUnmonitoredCount;
    emailTemplate.sheetErrors = sheetErrors;

    // Run code in template
    var htmlBody = emailTemplate.evaluate().getContent();

    // Finally send mail
    MailApp.sendEmail({
      to: "email@domain.com",
      subject: "ATLS Badge Status",
      body: null,
      htmlBody: htmlBody
    })
  }

}
