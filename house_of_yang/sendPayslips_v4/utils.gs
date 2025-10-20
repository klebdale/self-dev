/**
 * Generic Utility Functions. These functions should be pure or only handle 
 * low-level I/O like fetching raw data or writing simple values.
 */

// ====================================================================
// A. DATA I/O UTILITIES
// ====================================================================

/**
 * Creates a map of full employee name to nickname from the SALARY sheet.
 * @returns {Object} A map { 'Full Name': 'Nickname' }
 */
function getNameToNicknameMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName(SHEET_NAMES.SALARY);
  if (!salarySheet) throw new Error(`Sheet "${SHEET_NAMES.SALARY}" not found.`);

  const config = SALARY_CONFIG;
  const startRow = config.MAP_START_ROW;
  const startCol = config.EMPLOYEE_NAME_COL;
  
  // Fetch columns B and C (Name and Nickname)
  const name_nickname_data = salarySheet
    .getRange(startRow, startCol, salarySheet.getLastRow() - startRow + 1, 2)
    .getValues()
    .filter(row => row[0] && row[1]); 

  // Create map: { 'Full Name': 'Nickname' }
  return Object.fromEntries(name_nickname_data.map(row => [
    row[0].toString().trim(), 
    row[1].toString().trim()
  ]));
}


// ====================================================================
// B. SHEET MANIPULATION UTILITIES
// ====================================================================

/**
 * Clears the deduction and penalty cells on a given employee sheet.
 * @param {string} sheetName - The name of the employee payslip sheet (nickname).
 */
function clearPayslipDataCells(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const { COL_DATE, COL_ITEM, COL_AMOUNT, ROW_OTHER_DEDUCTIONS, ROW_PENALTIES } = PAYSLIP_OUTPUT;
  
  // Clear Other Deductions (Row 20)
  sheet.getRange(ROW_OTHER_DEDUCTIONS, COL_DATE, 1, 3).clearContent(); // 3 columns wide

  // If the columns are separated, we could do this:
  // sheet.getRange(ROW_OTHER_DEDUCTIONS, COL_DATE, 1, 3).clearContent();
  // sheet.getRange(ROW_OTHER_DEDUCTIONS, COL_ITEM, 1, 3).clearContent();
  // sheet.getRange(ROW_OTHER_DEDUCTIONS, COL_AMOUNT, 1, 3).clearContent();
  
  // Clear Penalties (Row 21)
  sheet.getRange(ROW_PENALTIES, COL_DATE, 1, 3).clearContent(); // 3 columns wide

  // If the columns are separated, we could do this:
  // sheet.getRange(ROW_PENALTIES, COL_DATE, 1, 3).clearContent();
  // sheet.getRange(ROW_PENALTIES, COL_ITEM, 1, 3).clearContent();
  // sheet.getRange(ROW_PENALTIES, COL_AMOUNT, 1, 3).clearContent();
}

/**
 * Writes the aggregated entries (Date, Item, Amount) to a single row using newline delimiters.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The employee payslip sheet.
 * @param {Array<Array<any>>} entries - An array of [Date, Item, Amount] entries.
 * @param {number} targetRow - The target row (ROW_OTHER_DEDUCTIONS or ROW_PENALTIES).
 */
function writePayslipEntry(sheet, entries, targetRow) {
  const { COL_DATE, COL_ITEM, COL_AMOUNT } = PAYSLIP_OUTPUT;

  if (entries.length > 0) {
    // Write aggregated values to single cells, separated by newlines
    sheet.getRange(targetRow, COL_DATE).setValue(entries.map(e => e[0]).join("\n")).setHorizontalAlignment("right");
    sheet.getRange(targetRow, COL_ITEM).setValue(entries.map(e => e[1]).join("\n"));
    sheet.getRange(targetRow, COL_AMOUNT).setValue(entries.map(e => e[2]).join("\n")).setHorizontalAlignment("right");
  }
}


/**
 * Finds the starting row indices of the Service Charge computation sections 
 * by checking only Column A.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The "SC Calc" sheet.
 * @returns {Object} { uptownRow: number, downtownRow: number } (1-based index)
 */
function findSCSections(sheet) {
  // Use config labels for search
  const UPTOWN_LABEL = SC_CONFIG.UPTOWN_LABEL;
  const DOWNTOWN_LABEL = SC_CONFIG.DOWNTOWN_LABEL;
  
  // Read all of column A, which is much faster than reading the entire sheet.
  const columnA = sheet.getRange("A:A").getValues();
  let uptownRow = null;
  let downtownRow = null;

  for (let r = 0; r < columnA.length; r++) {
    const cellValue = columnA[r][0];
    if (cellValue === UPTOWN_LABEL) uptownRow = r + 1;
    if (cellValue === DOWNTOWN_LABEL) downtownRow = r + 1;

    // Optimization: Stop searching once both are found.
    if (uptownRow && downtownRow) break; 
  }

  if (!uptownRow || !downtownRow) {
    throw new Error(`Could not find required section labels (${UPTOWN_LABEL}, ${DOWNTOWN_LABEL}) in SC Calc sheet.`);
  }

  return { uptownRow, downtownRow };
}

