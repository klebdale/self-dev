/**
 * Global Configuration and Constants for the Payroll App.
 */

// ====================================================================
// A. CORE SHEET NAMES
// ====================================================================

const SHEET_NAMES = {
  SALARY: "SALARY",
  SC_CALC: "SC Calc",
  ENTRY_LOG: "ENTRY LOG",
  OTHER_DEDUCTIONS: "OTHER DEDUCTIONS",
  EMAIL_LOG: "Payroll Email Log"
};

// Excluded sheets for the employee selection UI
const EXCLUDED_SHEETS = [
  SHEET_NAMES.SALARY, SHEET_NAMES.ENTRY_LOG, SHEET_NAMES.OTHER_DEDUCTIONS, SHEET_NAMES.SC_CALC,
  SHEET_NAMES.EMAIL_LOG, 'LABOR', 'Holidays', 'LEAVES', '13TH MONTH', 'SSS', 
  'Philhealth', 'PAGIBIG', 'PENALTIES', 'Benefits', 'OTHER BENEFITS', 
  'Scheduled Payslips', 'Uptown', 'Downtown', 'SC'
];


// ====================================================================
// B. DATA RANGES & CELL REFERENCES (SALARY SHEET)
// ====================================================================

const SALARY_CONFIG = {
  PAY_PERIOD_START_CELL: "J3",
  PAY_PERIOD_END_CELL: "L3",
  PAYMENT_DATE_CELL: "L2",
  // Columns for mapping Employee Name to Nickname
  MAP_START_ROW: 9, // Row where first employee name starts
  EMPLOYEE_NAME_COL: 2, // Column B
  NICKNAME_COL: 3,      // Column C
  SC_COLUMN_LETTER : "AA"
};


// ====================================================================
// C. PAYSLIP OUTPUT CONFIG (Employee Sheets)
// ====================================================================

// These define where the calculated items are written on the employee's payslip.
const PAYSLIP_OUTPUT = {
  // Input: Where employee email is found for mailing
  EMAIL_ADDRESS_CELL: "C4",
  EMPLOYEE_FULL_NAME_CELL: "C3",
  // Output: Hardcoded start cells for all written data - Computation of Other Deductions & Penalties
  COL_DATE: 3, // Column C
  COL_ITEM: 4, // Column D
  COL_AMOUNT: 5, // Column E
  // Output Rows:
  ROW_OTHER_DEDUCTIONS: 20,
  ROW_PENALTIES: 21,
};


// ====================================================================
// D. CALCULATION INPUT CONFIG (Deductions & Other Penalties Source Sheets)
// ====================================================================

// Column headers used for dynamic indexing (must match header text exactly)
const CALC_HEADERS = {
  DATE: "DATE",
  ITEM: "ITEM",
  COVERAGE: "COVERAGE",
  EMPLOYEE_PENALTY: "EMPLOYEE", // from ENTRY LOG - Column where employee names are found
  PENALTY_AMOUNT: "PENALTIES", // from ENTRY LOG - Penalty Amount to be deducted
  EMPLOYEE_DEDUCTION: "RESPONSIBLE", // from OTHER DEDUCTIONS - Employee name
  DEDUCTION_AMOUNT: "AMOUNT" // from OTHER DEDUCTIONS - Amount to be deducted
};


// ====================================================================
// E. SERVICE CHARGE CALCULATION CONFIG (SC Calc Sheet)
// ====================================================================

const SC_CONFIG = {
  // Maximum number of days in the pay period (used for column insertion)
  MAX_DAYS: 16,
  
  // Column index where dates and formulas are inserted (Column B)
  START_COLUMN: 2, 
  
  // Labels used to find the start rows for copying formulas
  UPTOWN_LABEL: "SERVICE CHARGE COMPUTATION - UPTOWN",
  DOWNTOWN_LABEL: "SERVICE CHARGE COMPUTATION - DOWNTOWN",
  
  // The row number in SC Calc that holds the SC PER HOUR rate for Uptown Employees
  SC_PER_HOUR_UPTOWN_ROW_NUM: 15,

  // The row number in SC Calc that holds the SC PER HOUR rate for Downtown Employees
  SC_PER_HOUR_DOWNTOWN_ROW_NUM: 40,
};