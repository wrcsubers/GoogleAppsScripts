// =========================================================================================================================
// ATLS Next Visit Reminder
// Runs at the beginning of every month and sends an email to team with the previous month's maintenance notes
// Written January 2023 by Cameron Woods - Cameron.Woods@Aristocrat.com
// =========================================================================================================================

function processNextVisitReminders() {

  // Variables from sheet
  // =========================================================================================================================
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const lastRow = spreadsheet.getLastRow();
  let firstRow = 2;
  // As more rows are added, this should ensure we're only scanning the last 50.  Otherwise it will slow down.
  if (lastRow > 50){
    firstRow = lastRow - 49;
  }

  var htmlTableContent = ""

  const dateNow = new Date();
  const dateLastMonthStart = new Date(dateNow.getFullYear(),dateNow.getMonth() - 1, 1, 0, 0, 0, 1)
  const dateLastMonthEnd = new Date(dateNow.getFullYear(),dateNow.getMonth(), 1, 0, 0, -1, 0)
  const month = dateNow.toLocaleString('default', { month: 'long' });

  let maintenanceArray = [];

  // Process Data from sheet
  // =========================================================================================================================
  
  // Fill array with relevant data based on date
  for (let i = firstRow; i <= lastRow; i++){
    let maintenanceDate = new Date(spreadsheet.getRange("C" + i).getValue());
    if (maintenanceDate > dateLastMonthStart && maintenanceDate < dateLastMonthEnd){
      let whoAreYou = spreadsheet.getRange("B" + i).getValue();
      let casinoName = spreadsheet.getRange("D" + i).getValue();
      let additionalWorkNeeded = spreadsheet.getRange("N" + i).getValue();
      maintenanceArray.push([casinoName,additionalWorkNeeded,whoAreYou])
    }
  }

  // Sort the array
  maintenanceArray.sort()

  // Generate an HTML Table from the array
  for (let i = 0; i < maintenanceArray.length; i++){
    htmlTableContent += '<tr><td>' + maintenanceArray[i][0] + '</td><td>' + maintenanceArray[i][1] + '</td><td>' + maintenanceArray[i][2] + '</td></tr>'
  }

  // Write Email and Send
  // =========================================================================================================================
  
  // Load email template in
  var emailTemplate = HtmlService.createTemplateFromFile("NextVisitReminder_EmailTemplate.html");

  // Assign variables in email template
  emailTemplate.month = month
  emailTemplate.htmlTableContent = htmlTableContent;

  // Run code in template
  var htmlBody = emailTemplate.evaluate().getContent();

  // Finally send mail
  MailApp.sendEmail({
    to: "email@domain.com",
    subject: "ATLS Maintenance Reminders",
    body: null,
    htmlBody: htmlBody
  })
  
}
