// =========================================================================================================================
// ATLS Maintenance Report Processor
// Runs after from submission and sends an email to team with details of each maintenance visit
// Written December 2023 by Cameron Woods - Cameron.Woods@Aristocrat.com
// =========================================================================================================================

function processMaintenanceReport() {

  // Variables from sheet
  // =========================================================================================================================

  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const lastRow = spreadsheet.getLastRow();

  const whoAreYou = spreadsheet.getRange("B" + lastRow).getValue();
  const maintenanceDate = spreadsheet.getRange("C" + lastRow).getValue();
  const casinoName = spreadsheet.getRange("D" + lastRow).getValue();
  const serversFaultLights = spreadsheet.getRange("E" + lastRow).getValue();
  const switchesFaultLights = spreadsheet.getRange("F" + lastRow).getValue();
  const sanFaultLights = spreadsheet.getRange("G" + lastRow).getValue();
  const upsWarnings = spreadsheet.getRange("H" + lastRow).getValue();
  const onelinkEGMSOnline = spreadsheet.getRange("J" + lastRow).getValue();
  const onelinkOtherErrors = spreadsheet.getRange("K" + lastRow).getValue();
  const dbLogsShrunk = spreadsheet.getRange("L" + lastRow).getValue();
  const additionalWorkDone = spreadsheet.getRange("M" + lastRow).getValue();
  const additionalWorkNeeded = spreadsheet.getRange("N" + lastRow).getValue();
  const additionalNotes = spreadsheet.getRange("O" + lastRow).getValue();
  const billingReports = spreadsheet.getRange("P" + lastRow).getValue();
  const uploadToShare = spreadsheet.getRange("Q" + lastRow).getValue();
  const reportPath = spreadsheet.getRange("R" + lastRow).getValue().toString().split('=')[1];

  // Trim the end name of the IRIS Report Attachment and set as attachment
  var report = DriveApp.getFileById(reportPath);
  if(report.getName().includes(' ')){
    report.setName(report.getName().split(' ')[0] + '.html');
  } 
  const emailAttachment = DriveApp.getFileById(reportPath).getAs("text/html");

  // Write Email and Send
  // =========================================================================================================================

  // Load email template in
  var emailTemplate = HtmlService.createTemplateFromFile("MaintenanceReport_EmailTemplate.html");

  // Assign variables in email template
  emailTemplate.whoAreYou = whoAreYou;
  emailTemplate.maintenanceDate = maintenanceDate;
  emailTemplate.casinoName = casinoName;
  emailTemplate.serversFaultLights = serversFaultLights;
  emailTemplate.switchesFaultLights = switchesFaultLights;
  emailTemplate.sanFaultLights = sanFaultLights;
  emailTemplate.upsWarnings = upsWarnings;
  emailTemplate.onelinkEGMSOnline = onelinkEGMSOnline;
  emailTemplate.onelinkOtherErrors = onelinkOtherErrors;
  emailTemplate.dbLogsShrunk = dbLogsShrunk;
  emailTemplate.additionalWorkDone = additionalWorkDone;
  emailTemplate.additionalWorkNeeded = additionalWorkNeeded;
  emailTemplate.additionalNotes = additionalNotes;
  emailTemplate.billingReports = billingReports;
  emailTemplate.uploadToShare = uploadToShare;

  // Run code in template
  var htmlBody = emailTemplate.evaluate().getContent();

  // Finally send mail
  MailApp.sendEmail({
    to: "email@domain.com",
    subject: "ATLS Maintenance Report - " + casinoName,
    body: null,
    htmlBody: htmlBody,
    attachments: [emailAttachment]
  })

}
