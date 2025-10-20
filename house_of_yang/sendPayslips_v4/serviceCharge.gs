/**
 * SERVICE CHARGE SERVICE: serviceCharge.gs
 * Main function to prepare the SC Calc sheet based on the pay period.
 * Adding of the Date Columns and copies the formulas from the previous dates to the SC Calc Sheet
 */
function updateServiceChargeComputation() {
  Logger.log("Starting Service Charge Computation Update...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(SHEET_NAMES.SC_CALC);
  const salarySheet = ss.getSheetByName(SHEET_NAMES.SALARY);
  
  if (!targetSheet) throw new Error(`Sheet "${SHEET_NAMES.SC_CALC}" not found.`);
  if (!salarySheet) throw new Error(`Sheet "${SHEET_NAMES.SALARY}" not found.`);
  
  const MAX_DAYS = SC_CONFIG.MAX_DAYS;
  const START_COL = SC_CONFIG.START_COLUMN;

  // 1️⃣ Get start and end dates (using config cells)
  const startDate = new Date(salarySheet.getRange(SALARY_CONFIG.PAY_PERIOD_START_CELL).getValue());
  const endDate = new Date(salarySheet.getRange(SALARY_CONFIG.PAY_PERIOD_END_CELL).getValue());
  
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    throw new Error("Invalid date found in SALARY sheet pay period cells.");
  }

  // 2️⃣ Insert new columns to clear space and push the static formula column
  targetSheet.insertColumnsBefore(START_COL, MAX_DAYS);

  // 3️⃣ Generate date values & count valid columns
  const dates = [];
  let currentDate = new Date(startDate);
  let validCols = 0;
  
  for (let i = 0; i < MAX_DAYS; i++) {
    if (currentDate <= endDate) {
      dates.push(new Date(currentDate));
      currentDate.setDate(currentDate.getDate() + 1);
      validCols++;
    } else {
      dates.push("");
    }
  }

  // 4️⃣ Find section start rows (using optimized utility from utils.gs)
  const { uptownRow, downtownRow } = findSCSections(targetSheet);

  // 5️⃣ Write only valid date columns
  targetSheet.getRange(uptownRow + 1, START_COL, 1, validCols).setValues([dates.slice(0, validCols)]);
  targetSheet.getRange(downtownRow + 1, START_COL, 1, validCols).setValues([dates.slice(0, validCols)]);

  // 6️⃣ Copy formulas only for valid date columns
  const sourceCol = START_COL + MAX_DAYS; 

  // --- UPTOWN Formulas ---
  const uptownEndRow = downtownRow - 1; 
  const uptownNumRows = uptownEndRow - (uptownRow + 1);
  if (uptownNumRows > 0) {
    const uptownSourceRange = targetSheet.getRange(uptownRow + 2, sourceCol, uptownNumRows, 1);
    const uptownTargetRange = targetSheet.getRange(uptownRow + 2, START_COL, uptownNumRows, validCols);
    uptownSourceRange.copyTo(uptownTargetRange, {contentsOnly: false});
  }


  // --- DOWNTOWN Formulas ---
  const lastRow = targetSheet.getLastRow();
  const downtownNumRows = lastRow - (downtownRow + 1);
  if (downtownNumRows > 0) {
    const downtownSourceRange = targetSheet.getRange(downtownRow + 2, sourceCol, downtownNumRows, 1);
    const downtownTargetRange = targetSheet.getRange(downtownRow + 2, START_COL, downtownNumRows, validCols);
    downtownSourceRange.copyTo(downtownTargetRange, {contentsOnly: false});
  }

  // 7️⃣ Call Macro to update SALARY sheet formula (The next step in the workflow)
  updateSalarySheetSCFormula(); 

  Logger.log("Service Charge computation updated successfully.");
}