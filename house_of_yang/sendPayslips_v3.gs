const EXCLUDED_SHEETS = [
  'LABOR', 'SALARY', 'ENTRY LOG', 'OTHER DEDUCTIONS', 'Holidays', 'LEAVES',
  '13TH MONTH', 'SSS', 'Philhealth', 'PAGIBIG', 'PENALTIES', 'Payroll Email Log',
  'Benefits', 'OTHER BENEFITS', 'Scheduled Payslips'
];


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Payslips')
    .addItem('Select Employees', 'checkboxEmployees')
    .addToUi();
}


function checkboxEmployees() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();


  // Get all sheets except excluded ones
  var employeeNames = sheets
    .map(sheet => sheet.getSheetName())
    .filter(name => !EXCLUDED_SHEETS.includes(name));

  // Create HTML form with checkboxes
  var html = `
    <form id="checkboxForm">
      <div>
        <button type="button" onclick="checkAll(true)">Check All</button>
        <button type="button" onclick="checkAll(false)">Uncheck All</button>
      </div>
      <br>
  `;

  employeeNames.forEach(name => {
    html += `<label><input type="checkbox" name="sheets" value="${name}" checked> ${name}</label><br>`;
  });

  html += `
      <br>
      <label><strong>Pick date and time to schedule:</strong><br>
        <input type="datetime-local" id="scheduledDatetime" required>
      </label>
      <br><br>

      <button onclick="scheduleWithDateTime()" type="button">
        📅 ✉️ Schedule Payslip Sending
      </button>
      <br><br>

      <button onclick="google.script.run.processPenaltiesAndSelectedSheets(getSelectedSheets()); google.script.host.close();" type="button">
        🧮 ✉️Calculate Deductions & Penalties THEN Send Payslip
      </button>
      <br><br>

      <button onclick="google.script.run.calculateDeductionsAndPenalties(); google.script.host.close();" type="button">
        🧮 Only Calculate Deductions & Penalties (All Employees)
      </button>
      <br><br>

      <button onclick="google.script.run.processSelectedSheets(getSelectedSheets()); google.script.host.close();" type="button">
        ✉️ Only Send Payslip (Selected Employees Only)
      </button>
      <br><br>
      
      <script>
        function getSelectedSheets() {
          const checkboxes = document.querySelectorAll('input[name="sheets"]:checked');
          return Array.from(checkboxes).map(cb => cb.value);
        }

        function checkAll(isChecked) {
          document.querySelectorAll('input[name="sheets"]').forEach(cb => cb.checked = isChecked);
        }

        function scheduleWithDateTime() {
          const selectedSheets = getSelectedSheets();
          const datetime = document.getElementById('scheduledDatetime').value;
          if (!datetime) {
            alert("Please pick a valid date and time.");
            return;
          }
          google.script.run.withSuccessHandler(() => {
            google.script.host.close();
          }).saveSchedule(selectedSheets, datetime);
        }
      </script>
    </form>
  `;

  // Display the form
  var ui = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Select Employees');
}


// A combination of functions that calculates deductions and penalties for ALL employees
// and sends payslips to SELECTED employees only
// runs after clicking "Calculate Deductions & Penalties THEN Send Payslip" of selected employees
function processPenaltiesAndSelectedSheets(selectedSheets) {
  if (!selectedSheets || selectedSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No employees selected.');
    return;
  }

  // clear the deduction and penalties cells from the selected sheets
  selectedSheets.forEach(name => clearPenaltyAndDeductionCells(name));
  
  // Calculate other deductions and penalties of ALL employees; but only
  // Pass the selected sheets to the sendPayslips function
  calculateOtherDeductions();
  calculatePenalties();
  Utilities.sleep(1500);
  sendPayslips(selectedSheets);
}


// function to handle calcualte deductions and penalties for ALL employees
function calculateDeductionsAndPenalties(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  // Get all sheets except excluded ones
  let employeeNames = sheets
    .map(sheet => sheet.getSheetName())
    .filter(name => !EXCLUDED_SHEETS.includes(name));

  employeeNames.forEach(name => clearPenaltyAndDeductionCells(name));

  calculateOtherDeductions();
  calculatePenalties();
}


// Function to handle the selected sheets
// runs after clicking "Only Send Payslip (Selected Employees Only)" button
function processSelectedSheets(selectedSheets) {
  if (!selectedSheets || selectedSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No employees selected.');
    return;
  }
  
  // Pass the selected sheets to the sendPayslips function
  sendPayslips(selectedSheets);
}


function sendPayslips(employeeNames = null) {
  //const ui = SpreadsheetApp.getUi(); // Get the UI for showing dialog boxes
  //ui.alert('Processing', 'The email sending process has started. Press "OK" and please wait...', ui.ButtonSet.OK);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  sheets.forEach(function(sheet) {
    if (employeeNames.includes(sheet.getSheetName())) {

      var name = sheet.getRange('C3').getValue();
      var email = sheet.getRange('C4').getValue();
      var payPeriodStart = new Date(sheet.getRange('C5').getValue());
      var payPeriodEnd = new Date(sheet.getRange('G5').getValue());
      var paymentDate = new Date(sheet.getRange('C23').getValue());

      var formattedPayPeriodStart = Utilities.formatDate(payPeriodStart, "GMT+8", 'MMM dd yyyy');
      var formattedPayPeriodEnd = Utilities.formatDate(payPeriodEnd, "GMT+8", 'MMM dd yyyy');
      var formattedPaymentDate = Utilities.formatDate(paymentDate, "GMT+8", 'MMM dd yyyy');

      var subject = 'Your Payroll Slip for ' + formattedPaymentDate;

      var message = "<p>Dear " + name + ",<br><br>I hope this email finds you well.<br><br>We are pleased to inform you that your payroll slip for " + formattedPayPeriodStart + " to " + formattedPayPeriodEnd + " has been generated.<br><br>Please review the attached payroll slip for a detailed breakdown. If you have any questions or notice any discrepancies, do not hesitate to reach out to the HR department.<br><br>Thank you for your continued hard work and dedication.<br><br>Best Regards,<br>House of Yang<br><br><i>Note: Please remember that your salary information is private and confidential</i>";
      
      //var pdf = convertSheetToPDF(sheet);
      var pdf = safeConvertSheetToPDF(sheet);

      if (email) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: message,
            attachments: [pdf]
          });
          Logger.log('Sent payslip to ' + email);
          logEmail(sheet.getSheetName(), email, 'Success');
        } catch (error) {
          Logger.log('Failed to send payslip to ' + email + ': ' + error.message);
          logEmail(sheet.getSheetName(), email, 'Failed: ' + error.message);
        }

        // Introduce a delay to avoid triggering rate limits
        Utilities.sleep(5000); // 5 seconds delay
        //var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
        //Logger.log("Remaining email quota: " + emailQuotaRemaining);
      }
    }
  });
  //ui.alert('Done', 'The email sending process has completed successfully.', ui.ButtonSet.OK);
}


function logEmail(sheetName, email, status) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = spreadsheet.getSheetByName('Payroll Email Log');

  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('Payroll Email Log');
    logSheet.appendRow(['Sheet Name', 'Email', 'Status', 'Date Sent']);
  }

  var dateSent = new Date();
  var newRow = [sheetName, email, status, Utilities.formatDate(dateSent, "GMT+8", 'MMM dd yyyy HH:mm:ss')];
  logSheet.insertRowBefore(2); // Insert a new row at position 2 (below the header)
  logSheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
}

function convertSheetToPDF(sheet) {
  var spreadsheet = sheet.getParent();
  var sheetId = sheet.getSheetId();
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?';

  var exportOptions = [
    'format=pdf',
    'size=letter',
    'portrait=true',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    'gid=' + sheetId
  ];

  var params = {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url + exportOptions.join('&'), params);
  return response.getBlob().setName(sheet.getName() + '.pdf');
}


// Just a wrapper function that allows exponential backoff retries for converting to PDF
// to avoid overwhelming api request
function safeConvertSheetToPDF(sheet, retries = 3) {
  for (var i = 0; i < retries; i++) {
    try {
      return convertSheetToPDF(sheet);
    } catch (e) {
      if (e.message.includes("Request failed")) {
        Logger.log("Retrying PDF export for " + sheet.getSheetName() + " in " + (Math.pow(2, i) * 1000) + " ms...");
        Utilities.sleep(Math.pow(2, i) * 1000); // Exponential backoff
      } else {
        throw e; // Rethrow other errors
      }
    }
  }
  throw new Error("PDF export failed after " + retries + " retries.");
}

// For Scheduled Send
function saveSchedule(selectedEmployees, datetimeString) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1500); // wait max 3 seconds
    const props = PropertiesService.getScriptProperties();

    // Store data
    props.setProperty("scheduledEmployees", JSON.stringify(selectedEmployees));
    props.setProperty("scheduledDateTime", datetimeString);

    // Schedule trigger
    scheduleTriggerForDateTime(datetimeString);
  } catch (e) {
    Logger.log("Could not acquire lock: " + e);
  } finally {
    lock.releaseLock();
  }
}


function scheduleTriggerForDateTime(datetimeString) {
  const date = new Date(datetimeString);
  if (isNaN(date.getTime())) {
    Logger.log("Invalid date.");
    return;
  }

  // Clear existing triggers of this type to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runScheduledPayslip') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("runScheduledPayslip")
    .timeBased()
    .at(date)
    .create();
}


function runScheduledPayslip() {
  const props = PropertiesService.getScriptProperties();
  const employeeList = JSON.parse(props.getProperty("scheduledEmployees") || "[]");

  if (employeeList.length > 0) {
    // Ask Shonu if he still manually ads some penalties, if not, use this
    //processPenaltiesAndSelectedSheets(employeeList)
    processSelectedSheets(employeeList);
  } else {
    Logger.log("No employees found to process.");
  }
}


function calculatePenalties() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const penalties_sheet = spreadsheet.getSheetByName("ENTRY LOG");
  const salary_sheet = spreadsheet.getSheetByName("SALARY");
  
  if (!penalties_sheet || !salary_sheet) {
    Logger.log("ENTRY LOG or SALARY sheets not found.");
    return;
  }

  const data = penalties_sheet.getDataRange().getValues(); 
  if (data.length === 0) return Logger.log("ENTRY LOG sheet is empty.");

  // Fetch EMPLOYEE & NICKNAME mapping (starting from row 7, column B)
  const startRowSalary = 7; 
  const startColSalary = 2; 
  const name_nickname_data = salary_sheet
    .getRange(startRowSalary, startColSalary, salary_sheet.getLastRow() - startRowSalary + 1, 2)
    .getValues()
    .filter(row => row[0] && row[1]); // Filter out empty rows

  let nameToNickname = Object.fromEntries(name_nickname_data.map(row => [row[0].toString().trim(), row[1].toString().trim()]));

  // Identify column indexes dynamically
  const headers = data[0]; 
  const dateCol = headers.indexOf("DATE");
  const itemCol = headers.indexOf("ITEM");
  const penaltyCol = headers.indexOf("PENALTIES");
  const responsibleCol = headers.indexOf("EMPLOYEE");
  const coverageCol = headers.indexOf("COVERAGE");

  if ([dateCol, itemCol, penaltyCol, responsibleCol, coverageCol].includes(-1)) {
    Logger.log("One or more required columns are missing.");
    return;
  }
  
  // Process the "ENTRY LOG" data
  let employeeData = {};
  for (let i = 1; i < data.length; i++) {
    // Skip if "COVERAGE" is not TRUE OR penalty is empty
    if (data[i][coverageCol] !== true || !data[i][penaltyCol]) continue;

    let responsible = data[i][responsibleCol];
    if (!responsible || !nameToNickname[responsible]) continue; // Skip if no matching nickname

    let formattedDate = data[i][dateCol] instanceof Date 
      ? Utilities.formatDate(data[i][dateCol], Session.getScriptTimeZone(), "MM-dd") 
      : data[i][dateCol];

    let entry = [formattedDate, data[i][itemCol], data[i][penaltyCol]];
    
    if (!employeeData[responsible]) employeeData[responsible] = [];
    employeeData[responsible].push(entry);
  }

  // Insert filtered data into each employee's sheet
  Object.entries(employeeData).forEach(([employee, entries]) => {
    let sheet = spreadsheet.getSheetByName(nameToNickname[employee]);
    if (!sheet) return Logger.log("Sheet for " + employee + " not found.");

    // Change when salary slip format changes
    // Row 19  = Penalties
    let row = 20, colDate = 3, colItem = 4, colAmount = 5;
    
    // Determine the number of rows to clear based on the longest entry list
    let numRows = entries.length || 1; 

    // Clear existing content in the target columns before writing new values
    // Safe to delete? clearPenaltyAndDeductionCells() exists already
    /*
    sheet.getRange(row, colDate, numRows, 1).clearContent();
    sheet.getRange(row, colItem, numRows, 1).clearContent();
    sheet.getRange(row, colAmount, numRows, 1).clearContent();
    Logger.log("Cleared: " + employee + "row: " + row + ", col:" + colDate + "," + colItem + "," + colAmount);
    */

    // Write all entries in just one cell, not new row every entry
    sheet.getRange(row, colDate).setValue(entries.map(e => e[0]).join("\n")).setHorizontalAlignment("right");
    sheet.getRange(row, colItem).setValue(entries.map(e => e[1]).join("\n"));
    sheet.getRange(row, colAmount).setValue(entries.map(e => e[2]).join("\n")).setHorizontalAlignment("right");;
  });

  Logger.log("Penalties calculated successfully.");
}


function calculateOtherDeductions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const brkg_splg_sheet = spreadsheet.getSheetByName("OTHER DEDUCTIONS");
  const salary_sheet = spreadsheet.getSheetByName("SALARY");
  
  if (!brkg_splg_sheet || !salary_sheet) {
    Logger.log("OTHER DEDUCTIONS or SALARY sheets not found.");
    return;
  }

  const data = brkg_splg_sheet.getDataRange().getValues(); 
  
  if (data.length === 0) return Logger.log("OTHER DEDUCTIONS sheet is empty.");

  // Identify column indexes dynamically
  const headers = data[0]; 
  const dateCol = headers.indexOf("DATE");
  const itemCol = headers.indexOf("ITEM");
  const amountCol = headers.indexOf("AMOUNT");
  const responsibleCol = headers.indexOf("RESPONSIBLE");
  const coverageCol = headers.indexOf("COVERAGE");

  if ([dateCol, itemCol, amountCol, responsibleCol, coverageCol].includes(-1)) {
    Logger.log("One or more required columns are missing.");
    return;
  }

  // Fetch EMPLOYEE & NICKNAME mapping (starting from row 7, column B)
  const startRowSalary = 7; 
  const startColSalary = 2; 
  const name_nickname_data = salary_sheet
    .getRange(startRowSalary, startColSalary, salary_sheet.getLastRow() - startRowSalary + 1, 2)
    .getValues()
    .filter(row => row[0] && row[1]); // Filter out empty rows

  let nameToNickname = Object.fromEntries(name_nickname_data.map(row => [row[0].toString().trim(), row[1].toString().trim()]));

  let employeeData = {};

  // Process the "OTHER DEDUCTIONS" data
  for (let i = 1; i < data.length; i++) {
    if (data[i][coverageCol] !== true) continue; // Skip if "COVERAGE" is not TRUE

    // if TRUE, continue here
    let responsible = data[i][responsibleCol];
    if (!responsible || !nameToNickname[responsible]) continue; // Skip if no matching nickname
    
    // Format the date
    let formattedDate = data[i][dateCol] instanceof Date 
      ? Utilities.formatDate(data[i][dateCol], Session.getScriptTimeZone(), "MM-dd") 
      : data[i][dateCol];

    // Format entry to be placed in employee salary sheet
    let entry = [formattedDate, data[i][itemCol], data[i][amountCol]];
    
    if (!employeeData[responsible]) employeeData[responsible] = [];
    employeeData[responsible].push(entry);
  }

  // Insert filtered data entry into each employee's sheet
  Object.entries(employeeData).forEach(([employee, entries]) => {
    let sheet = spreadsheet.getSheetByName(nameToNickname[employee]);
    if (!sheet) return Logger.log("Sheet for " + employee + " not found.");

    // Hardcoded values for the salary slip to insert the entries.
    // Change when salary slip format changes
    // Row 19  = Other Deduction
    let row = 19, colDate = 3, colItem = 4, colAmount = 5;
  
    // Write all entries in just one cell, not new row every entry
    sheet.getRange(row, colDate).setValue(entries.map(e => e[0]).join("\n")).setHorizontalAlignment("right");
    sheet.getRange(row, colItem).setValue(entries.map(e => e[1]).join("\n"));
    sheet.getRange(row, colAmount).setValue(entries.map(e => e[2]).join("\n")).setHorizontalAlignment("right");
  });

  Logger.log("Other Deductions calculated successfully.");
}


function clearPenaltyAndDeductionCells(sheetName){
  // for Other Deductions and Entry Log/Penalties
  // Change when salary slip format changes
  // Row 19  = Other Deduction, Row 20 = Penalties
  let rowOD = 19, rowP = 20, colDate = 3, colItem = 4, colAmount = 5;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Clear existing content in the target columns before writing new values
  sheet.getRange(rowOD, colDate, 1, 1).clearContent();
  sheet.getRange(rowOD, colItem, 1, 1).clearContent();
  sheet.getRange(rowOD, colAmount, 1, 1).clearContent();

  sheet.getRange(rowP, colDate, 1, 1).clearContent();
  sheet.getRange(rowP, colItem, 1, 1).clearContent();
  sheet.getRange(rowP, colAmount, 1, 1).clearContent();

  Logger.log("Cleared: " + sheetName + "rows: " + rowP + "," + rowOD + "col:" + colDate + "," + colItem + "," + colAmount);
}
