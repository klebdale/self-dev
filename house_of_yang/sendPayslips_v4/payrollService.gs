/**
 * Core Payroll Service Layer. Contains the business logic for calculating 
 * and applying deductions/penalties.
 */

/**
 * Calculates and applies items (deductions or penalties) from a source sheet
 * to employee payslip sheets. This is the new, consolidated core function.
 * * @param {string} sourceSheetName - The name of the source sheet (e.g., 'ENTRY LOG').
 * @param {string} employeeHeader - The column header for the employee responsible.
 * @param {string} amountHeader - The column header for the amount (PENALTIES/AMOUNT).
 * @param {number} targetRow - The destination row in the payslip (ROW_PENALTIES/ROW_OTHER_DEDUCTIONS).
 * @param {Array<string>} [selectedNicknames] - Optional list of employee sheets to process.
 */
function calculateAndApplyItems(sourceSheetName, employeeHeader, amountHeader, targetRow, selectedNicknames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) return Logger.log(`${sourceSheetName} sheet not found.`);

  const nameToNickname = getNameToNicknameMap();
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) return Logger.log(`${sourceSheetName} sheet is empty.`);

  // Identify column indexes dynamically using config headers
  const headers = data[0]; 
  const dateCol = headers.indexOf(CALC_HEADERS.DATE);
  const itemCol = headers.indexOf(CALC_HEADERS.ITEM);
  const amountCol = headers.indexOf(amountHeader);
  const responsibleCol = headers.indexOf(employeeHeader);
  const coverageCol = headers.indexOf(CALC_HEADERS.COVERAGE);
  
  // Basic validation
  if ([dateCol, itemCol, amountCol, responsibleCol, coverageCol].includes(-1)) {
    return Logger.log(`One or more required columns are missing in ${sourceSheetName}.`);
  }

  let employeeData = {};
  
  // 1. DATA EXTRACTION AND AGGREGATION
  for (let i = 1; i < data.length; i++) {
    // Filter logic: Must be TRUE for COVERAGE and must have an amount
    if (data[i][coverageCol] !== true || !data[i][amountCol]) continue;

    const responsible = data[i][responsibleCol];
    const nickname = nameToNickname[responsible];
    
    // Filter by selected nicknames if provided
    if (!responsible || !nickname || (selectedNicknames && !selectedNicknames.includes(nickname))) continue;
    
    // Format the date
    const formattedDate = data[i][dateCol] instanceof Date 
      ? Utilities.formatDate(data[i][dateCol], Session.getScriptTimeZone(), "MM-dd")
      : data[i][dateCol];

    const entry = [formattedDate, data[i][itemCol], data[i][amountCol]];
    
    if (!employeeData[responsible]) employeeData[responsible] = [];
    employeeData[responsible].push(entry);
  }

  // 2. DATA APPLICATION (The only I/O write in this service)
  Object.entries(employeeData).forEach(([employee, entries]) => {
    let payslipSheet = ss.getSheetByName(nameToNickname[employee]);
    if (!payslipSheet) return Logger.log(`Sheet for ${employee} not found.`);

    writePayslipEntry(payslipSheet, entries, targetRow);
  });
  
  Logger.log(`${sourceSheetName} calculation applied successfully.`);
}


// ====================================================================
// PUBLIC SERVICE CALLS (For Controller Layer)
// ====================================================================

/**
 * Public function for calculating and applying Other Deductions.
 * @param {Array<string>} [selectedNicknames] - Optional list of employee sheets to process.
 */
function calculateAndApplyOtherDeductions(selectedNicknames) {
  calculateAndApplyItems(
    SHEET_NAMES.OTHER_DEDUCTIONS,
    CALC_HEADERS.EMPLOYEE_DEDUCTION,
    CALC_HEADERS.DEDUCTION_AMOUNT,
    PAYSLIP_OUTPUT.ROW_OTHER_DEDUCTIONS,
    selectedNicknames
  );
}

/**
 * Public function for calculating and applying Penalties.
 * @param {Array<string>} [selectedNicknames] - Optional list of employee sheets to process.
 */
function calculateAndApplyPenalties(selectedNicknames) {
  calculateAndApplyItems(
    SHEET_NAMES.ENTRY_LOG,
    CALC_HEADERS.EMPLOYEE_PENALTY,
    CALC_HEADERS.PENALTY_AMOUNT,
    PAYSLIP_OUTPUT.ROW_PENALTIES,
    selectedNicknames
  );
}