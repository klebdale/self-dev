/**
 * CONTROLLER LAYER: sendEmailPayslip.gs
 *
 * This file handles User Interface (UI), trigger scheduling, and orchestrates
 * the main workflow by calling functions from the Service and Utility layers.
 * It uses constants defined in config.gs.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Payslips')
    .addItem('Select Employees', 'checkboxEmployees')
    .addItem('Calculate SC', 'addDateColumns') // SC Calculation remains here for now
    .addItem('Calculate All (Deductions/SC)', 'calculateDeductionsAndPenalties')
    .addToUi();
}


function checkboxEmployees() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  // Get all sheets except excluded ones (uses EXCLUDED_SHEETS from config.gs)
  const employeeNames = sheets
    .map(sheet => sheet.getSheetName())
    .filter(name => !EXCLUDED_SHEETS.includes(name));

  // Create HTML form for employee selection
  const html = `
    <form id="checkboxForm">
      <div>
        <button type="button" onclick="checkAll(true)">Check All</button>
        <button type="button" onclick="checkAll(false)">Uncheck All</button>
      </div>
      <br>
      <div id="employeeList">
  ` +
  employeeNames.map(name => 
    `<label><input type="checkbox" name="sheets" value="${name}" checked> ${name}</label><br>`
  ).join('') + 
  `
      </div>
      <br>
      <label><strong>Pick date and time to schedule (Optional):</strong><br>
        <input type="datetime-local" id="scheduleTime"><br>
        <small>Schedules the full calculation and send process.</small>
      </label>
      <br><br>
      
      <input type="button" value="Calculate & Send Payslips" id="calcAndSendBtn" onclick="runScript('calculateAndSendPayslips', 'calcAndSendBtn')" />
      <input type="button" value="Only Send Payslips" id="sendOnlyBtn" onclick="runScript('onlySendPayslip', 'sendOnlyBtn')" />
      <input type="button" value="Schedule Payslips" id="scheduleBtn" onclick="schedulePayslips(getSelectedSheets(), document.getElementById('scheduleTime').value)" />
    </form>


    <script>
      function getSelectedSheets() {
        return Array.from(document.querySelectorAll('input[name="sheets"]:checked')).map(el => el.value);
      }


      function checkAll(checked) {
        document.querySelectorAll('input[name="sheets"]').forEach(el => el.checked = checked);
      }


      function runScript(serverFunctionName, buttonId) {
        const selectedSheets = getSelectedSheets();
        if (selectedSheets.length === 0) {
            alert('Please select employees.');
            return;
        }

        const button = document.getElementById(buttonId);
        
        // 1. START: Disable button and show loading text
        button.disabled = true;
        const originalValue = button.value;
        button.value = 'Processing... Please wait.';

        // 2. Call server-side function with handlers
        google.script.run
          .withSuccessHandler(function(e) {
            // SUCCESS: Re-enable button and show completion message
            button.disabled = false;
            button.value = originalValue;
            alert('Operation complete! Check the log sheet.');
          })
          .withFailureHandler(function(e) {
            // FAILURE: Re-enable button and show error message
            button.disabled = false;
            button.value = originalValue;
            alert('An error occurred: ' + e.message);
            console.error(e);
          })
          [serverFunctionName](selectedSheets); // Dynamically call the correct function
      }


      function schedulePayslips(selectedSheets, datetimeString) {
        if (!datetimeString) {
          alert('Please select a date and time for scheduling.');
          return;
        }
        if (selectedSheets.length === 0) {
          alert('Please select employees to schedule.');
          return;
        }
        google.script.run.withSuccessHandler(() => {
          google.script.host.close();
        }).saveSchedule(selectedSheets, datetimeString);
      }
    </script>
  `;

  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput(html)
    .setTitle('Payslip Selection & Control')
    .setWidth(300));
}


// -------------------------------------------------------------
// CONTROLLER WORKFLOW FUNCTIONS
// -------------------------------------------------------------

/**
 * Calculates penalties/deductions, applies the SC formula, and sends payslips 
 * only for selected employees. (The main "Run" button)
 * @param {Array<string>} selectedNicknames - List of employee sheets (nicknames).
 */
function calculateAndSendPayslips(selectedNicknames) {
  if (!selectedNicknames || selectedNicknames.length === 0) {
    SpreadsheetApp.getUi().alert('No employees selected.');
    return;
  }
  
  // 1. Clear cells for selected employees (Utility Layer)
  selectedNicknames.forEach(name => clearPayslipDataCells(name));

  // 2. Run Service Charge calculation (Service Charge file)
  updateServiceChargeComputation();

  // 3. Calculate and apply deductions and penalties (Service Layer)
  calculateAndApplyOtherDeductions(selectedNicknames); 
  calculateAndApplyPenalties(selectedNicknames); 
  
  Utilities.sleep(1500);

  // 4. Send Payslips (Controller)
  sendPayslips(selectedNicknames); 
}


/**
 * Runs the full calculation and application process for ALL employees.
 * (The "Calculate All" menu item)
 */
function calculateDeductionsAndPenalties(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get ALL employee nicknames for clearing
  const employeeNicknames = ss.getSheets()
    .map(sheet => sheet.getSheetName())
    .filter(name => !EXCLUDED_SHEETS.includes(name));
    
  // 2. Clear cells for ALL employee sheets (Utility Layer)
  employeeNicknames.forEach(name => clearPayslipDataCells(name));

  // 3. Run Service Charge calculation (Service Charge file)
  updateServiceChargeComputation()
  
  // 4. Calculate and apply deductions and penalties for ALL (Service Layer)
  calculateAndApplyOtherDeductions(); // No argument = ALL
  calculateAndApplyPenalties();       // No argument = ALL
  
  SpreadsheetApp.getUi().alert("Calculation Complete", "All deductions, penalties, and Service Charges have been calculated and applied.", SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * Sends payslips only for selected employees, assuming calculation is current.
 * @param {Array<string>} selectedNicknames - List of employee sheets (nicknames).
 */
function onlySendPayslip(selectedNicknames) {
  if (!selectedNicknames || selectedNicknames.length === 0) {
    SpreadsheetApp.getUi().alert('No employees selected.');
    return;
  }
  Logger.log("Selected Sheets: " + selectedNicknames.join(', '));
  
  // NOTE: SC, Deduction, and Penalties are NOT calculated here, as per "Send Only" request.
  sendPayslips(selectedNicknames);
}


/**
 * Processes a list of employee payslip sheet names (nicknames),
 * converts each sheet to a PDF, and emails it to the employee.
 * @param {Array<string>} employeeNicknames - A list of employee sheet names (nicknames).
 */
function sendPayslips(employeeNicknames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Get configuration and core details (using config.gs)
  const salarySheet = ss.getSheetByName(SHEET_NAMES.SALARY);
  if (!salarySheet) {
    ui.alert("Error", "SALARY sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  // Get pay period and date from defined config cells
  const PAY_PERIOD_START_RAW = salarySheet.getRange(SALARY_CONFIG.PAY_PERIOD_START_CELL).getValue();
  const PAY_PERIOD_END_RAW = salarySheet.getRange(SALARY_CONFIG.PAY_PERIOD_END_CELL).getValue();
  const PAYMENT_DATE_RAW = salarySheet.getRange(SALARY_CONFIG.PAYMENT_DATE_CELL).getValue();

  if (!PAY_PERIOD_START_RAW || !PAY_PERIOD_END_RAW || !PAYMENT_DATE_RAW) {
    ui.alert("Error", "Pay Period Start, Pay Period End, Payment Date not set in SALARY sheet.", ui.ButtonSet.OK);
    return;
  }
  
  // Format dates for email subject/body
  const formattedPayPeriodStart = Utilities.formatDate(new Date(PAY_PERIOD_START_RAW), Session.getScriptTimeZone(), "MMM dd");
  const formattedPayPeriodEnd = Utilities.formatDate(new Date(PAY_PERIOD_END_RAW), Session.getScriptTimeZone(), "MMM dd");
  const formattedPaymentDate = Utilities.formatDate(new Date(PAYMENT_DATE_RAW), Session.getScriptTimeZone(), "MMMM dd, yyyy");
  
  const EMAIL_SUBJECT = `Your Payslip for ${formattedPaymentDate}`;

  // 2. Iterate through selected employees and process
  employeeNicknames.forEach(sheetName => {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Skipping: Sheet named '${sheetName}' not found.`);
        return;
      }
      
      // Get employee's email address from config cell
      const employeeEmail = sheet.getRange(PAYSLIP_OUTPUT.EMAIL_ADDRESS_CELL).getValue(); 
      const employeeName = sheet.getRange(PAYSLIP_OUTPUT.EMPLOYEE_FULL_NAME_CELL).getValue(); 

      const EMAIL_BODY = `
        <html>
          <body>
            <p>Dear ${employeeName},</p>
            
            <p>I hope this email finds you well.</p>
            
            <p>
              We are pleased to inform you that your payroll slip for ${formattedPayPeriodStart} to ${formattedPayPeriodEnd} has been generated.
              <br><br>
              Please review the attached payroll slip for a detailed breakdown. If you have any questions or notice any discrepancies, do not hesitate to reach out to the HR department.
            </p>
            
            <p>
              Thank you for your continued hard work and dedication.
            </p>
            
            <p>
              Best Regards,
              <br>
              House of Yang
            </p>
            
            <p style="font-size: 0.9em; color: #666;">
              <i>Note: Please remember that your salary information is private and confidential.</i>
            </p>
          </body>
        </html>
    `;

      if (!employeeEmail) {
        Logger.log(`Skipping: ${employeeName} due to missing email.`);
        logEmail(employeeName, "N/A", formattedPaymentDate, "SKIPPED", "No email address found in payslip.");
        return;
      }
      
      if (MailApp.getRemainingDailyQuota() <= 0) {
        ui.alert("Warning", "Daily email quota reached. Stopping process.", ui.ButtonSet.OK);
        return; 
      }

      // 3. Convert the sheet to PDF (Robust Utility)
      const pdfBlob = safeConvertSheetToPDF(sheet);

      // 4. Send the email
      MailApp.sendEmail({
        to: employeeEmail,
        subject: EMAIL_SUBJECT,
        htmlBody: EMAIL_BODY,
        attachments: [pdfBlob.setName(`Payslip - ${employeeName} - ${formattedPaymentDate}.pdf`)]
      });

      // 5. Log the outcome
      logEmail(employeeName, employeeEmail, formattedPaymentDate, "SUCCESS");

      // Rate Limit Mitigation: Pause to respect Google's MailApp limits
      Utilities.sleep(5000); // 5 seconds wait per email

    } catch (e) {
      Logger.log(`Failed to send payslip for ${sheetName}. Error: ${e.toString()}`);
      logEmail(sheetName, "N/A", formattedPaymentDate, "FAILED", e.toString());
      // Continue to the next employee even if one fails
      Utilities.sleep(1000); 
    }
  });

  ui.alert("Success", `Payslips sent for ${employeeNicknames.length} employees. Check the '${SHEET_NAMES.EMAIL_LOG}' for details.`, ui.ButtonSet.OK);
}


// -------------------------------------------------------------
// UTILITY FUNCTIONS (Specific to Email/PDF/Logging)
// -------------------------------------------------------------

/**
 * Logs the result of an email attempt to the audit sheet.
 */
function logEmail(employeeName, employeeEmail, payPeriod, status, details = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.EMAIL_LOG);
  if (!logSheet) {
    Logger.log(`ERROR: Email Log sheet '${SHEET_NAMES.EMAIL_LOG}' not found.`);
    return;
  }

  // Define the log entry array
  const logEntry = [
    employeeName,
    employeeEmail,
    payPeriod,
    status,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy HH:mm:ss"),
    details
  ];

  logSheet.insertRowBefore(2);
  logSheet.getRange(2, 1, 1, logEntry.length).setValues([logEntry]);
}


/**
 * Attempts to convert a sheet to PDF with exponential backoff on failure.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to convert.
 * @returns {GoogleAppsScript.Base.Blob} The PDF blob.
 */
function safeConvertSheetToPDF(sheet) {
  const MAX_RETRIES = 5;
  for (let i = 0; i < MAX_RETRIES; i++) {
    try {
      return convertSheetToPDF(sheet);
    } catch (e) {
      Logger.log(`PDF conversion failed (Attempt ${i + 1}/${MAX_RETRIES}): ${e}`);
      if (i === MAX_RETRIES - 1) {
        throw new Error(`Failed to convert sheet to PDF after ${MAX_RETRIES} attempts.`);
      }
      // Exponential backoff (1s, 2s, 4s, 8s, ...)
      Utilities.sleep(Math.pow(2, i) * 1000);
    }
  }
}

/**
 * Converts a Google Sheet into a PDF blob using UrlFetchApp.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to convert.
 * @returns {GoogleAppsScript.Base.Blob} The PDF blob.
 */
function convertSheetToPDF(sheet) {
  const ss = sheet.getParent();
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  // Reference for PDF export parameters:
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${sheetId}` +
    // These parameters control the appearance of the exported PDF:
    '&size=A4' + // paper size
    '&portrait=true' + // portrait orientation
    '&fitw=true' + // fit to width
    '&sheetnames=false&printtitle=false&pagenumbers=false' + // exclude UI elements
    '&gridlines=false&fzr=false'; // exclude gridlines and freezing rows

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
      throw new Error(`PDF export failed with status code ${response.getResponseCode()}`);
  }

  return response.getBlob();
}


// -------------------------------------------------------------
// SCHEDULING FUNCTIONS
// -------------------------------------------------------------

function saveSchedule(selectedEmployees, datetimeString) {
  const lock = LockService.getScriptLock();
  try {
    // Wait for the lock to be acquired
    lock.waitLock(3000); 
    const props = PropertiesService.getScriptProperties();

    // Store data (using lock ensures atomic operation)
    props.setProperty("scheduledEmployees", JSON.stringify(selectedEmployees));
    props.setProperty("scheduledDateTime", datetimeString);

    // Schedule trigger
    scheduleTriggerForDateTime(datetimeString);
  } catch (e) {
    Logger.log("Could not acquire lock or scheduling failed: " + e);
  } finally {
    lock.releaseLock();
  }
}


function scheduleTriggerForDateTime(datetimeString) {
  const date = new Date(datetimeString);
  if (isNaN(date.getTime())) {
    Logger.log("Invalid date provided for scheduling.");
    return;
  }

  // Clear existing triggers of this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runScheduledPayslip') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create the new time-based trigger
  ScriptApp.newTrigger("runScheduledPayslip")
    .timeBased()
    .at(date)
    .create();
    
  Logger.log(`Scheduled payslip run for: ${date.toString()}`);
}


function runScheduledPayslip() {
  const props = PropertiesService.getScriptProperties();
  const employeeList = JSON.parse(props.getProperty("scheduledEmployees") || "[]");

  if (employeeList.length > 0) {
    // Scheduled runs MUST perform the full calculation and send process
    calculateAndSendPayslips(employeeList); 
  } else {
    Logger.log("No employees found to process for scheduled run.");
  }
}