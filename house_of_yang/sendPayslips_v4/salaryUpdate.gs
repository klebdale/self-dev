/**
 * SALARY UPDATE SERVICE: salaryUpdate.gs
 * Contains logic to update formulas in the main SALARY sheet.
 */

/**
 * Updates the Service Charge computation formula for all employees 
 * in the SALARY sheet using non-UI methods and conditional logic 
 * for Uptown/Downtown SC rates.
 */
/**
 * SALARY UPDATE SERVICE: salaryUpdate.gs
 * Contains logic to update formulas in the main SALARY sheet.
 */
function updateSalarySheetSCFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName(SHEET_NAMES.SALARY);
  const SC_SHEET_NAME = SHEET_NAMES.SC_CALC;
  const SC_CalcSheet = ss.getSheetByName(SC_SHEET_NAME);

  if (!salarySheet || !SC_CalcSheet) throw new Error("Required sheets not found.");

  // --- 1. Get Dynamic Configuration ---
  const { MAX_DAYS, SC_PER_HOUR_UPTOWN_ROW_NUM, SC_PER_HOUR_DOWNTOWN_ROW_NUM } = SC_CONFIG;
  const START_ROW_SALARY_CONFIG = SALARY_CONFIG.MAP_START_ROW; 
  const SC_COLUMN_LETTER = "AA"; 
  
  // Find sectional start rows (boundary line for the conditional formula)
  const { downtownRow } = findSCSections(SC_CalcSheet); // Only need downtownRow for the boundary
  
  // Determine the dynamic end column letter (e.g., 'Q' for 16 days)
  const SC_START_COL_INDEX = SC_CONFIG.START_COLUMN;
  const SC_END_COL_INDEX = SC_START_COL_INDEX + MAX_DAYS - 1; 
  const SC_END_COL_LETTER = columnToLetter(SC_END_COL_INDEX);

  // --- 2. Determine the Target Range ---
  const columnBValues = salarySheet.getRange("B:B").getValues();
  let lastRow = columnBValues.length;
  while (lastRow > 0 && !columnBValues[lastRow - 1][0]) {
    lastRow--;
  }

  // Define the target range (e.g., AA7:AALastRow)
  const targetRange = salarySheet.getRange(SC_COLUMN_LETTER + START_ROW_SALARY_CONFIG + ":" + SC_COLUMN_LETTER + lastRow);

  // Get range metrics for the loop
  const startRow = targetRange.getRow();
  const numRows = targetRange.getNumRows();

  const formulas = [];

  // --- 3. FIX: Dynamically Build the Formula Array for Each Row ---
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i; 
    const employeeRef = 'B' + currentRow; 

    // 1. Employee's match row in SC Calc
    const MATCH_EMPLOYEE_ROW = `MATCH(${employeeRef}, '${SC_SHEET_NAME}'!A:A, 0)`;

    // 2. Conditional SC Rate Row Number (e.g., 15 or 40)
    const SC_RATE_ROW_CONDITIONAL = `IF(${MATCH_EMPLOYEE_ROW} < ${downtownRow}, ${SC_PER_HOUR_UPTOWN_ROW_NUM}, ${SC_PER_HOUR_DOWNTOWN_ROW_NUM})`;

    // 3. Construct the INDIRECT string argument.
    // Syntax must be: INDIRECT("'SC Calc'!" & ADDRESS(row, col) & ":" & ADDRESS(row, col))
    const INDIRECT_ARG = 
      // Literal sheet name: 'SC Calc'!
      `"'${SC_SHEET_NAME}'!" & ` +
      
      // Start Address: B15 or B40 (col 2)
      `ADDRESS(${SC_RATE_ROW_CONDITIONAL}, ${SC_START_COL_INDEX}, 4, TRUE) & ":" & ` +
      
      // End Address: Q15 or Q40 (col 17)
      `ADDRESS(${SC_RATE_ROW_CONDITIONAL}, ${SC_END_COL_INDEX}, 4, TRUE)`;


    // Core SUMPRODUCT Formula
    const SUMPRODUCT_FORMULA = `SUMPRODUCT(
      INDEX('${SC_SHEET_NAME}'!$B:$${SC_END_COL_LETTER}, ${MATCH_EMPLOYEE_ROW}, 0),
      INDIRECT(${INDIRECT_ARG})
    )`;

    // Final Formula: Wrap SUMPRODUCT in IFERROR and return "-" on error
    const formulaString = `=IFERROR(${SUMPRODUCT_FORMULA}, "-")`;

    // Add the unique formula string to the array
    formulas.push([formulaString]);
  }

  // --- 4. Apply Formulas ---
  targetRange.setFormulas(formulas);
  
  Logger.log(`Conditional SC formulas updated in ${SC_COLUMN_LETTER}${startRow}:${SC_COLUMN_LETTER}${lastRow}.`);
}


/**
 * Helper function to convert column index to letter (e.g., 17 -> Q)
 * @param {number} colIndex 1-based column index
 * @returns {string} Column letter
 */
function columnToLetter(colIndex) {
  let temp, letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = Math.floor((colIndex - 1) / 26);
  }
  return letter;
}