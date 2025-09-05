// --- CONFIGURATION ---
var MASTER_SPREADSHEET_ID = '161uw5s1lOwhV7YTX8uLKDw6TMl_QLEpSqxmHEftZzVA'; // User-provided Master Sheet ID
var CHECKER_SHEET_NAME = 'Checker';
var LOG_SHEET_NAME = 'Log';
var HELPER_SHEET_NAME = '_QueryHelper'; // Hidden sheet for temporary formulas
var NEW_ROW_HIGHLIGHT_COLOR = '#ffcccc'; // Light red for duplicate rows

// --- UI & MENU FUNCTIONS ---

/**
 * Adds a custom menu and automatically opens the sidebar when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BHG DeDuper')
    .addItem('Open Sidebar', 'showDuplicateCheckerSidebar')
    .addToUi();
  showDuplicateCheckerSidebar(); // Automatically open the sidebar
}


/**
 * Serves the main HTML file to the sidebar.
 */
function showDuplicateCheckerSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BHG DeDuper')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}


// --- API FUNCTIONS (Called from Frontend) ---

/**
 * Calculates and returns duplicate health statistics since the previous Monday.
 * @return {object} An object containing percentage, totalChecked, and totalDuplicates.
 */
function getDuplicateHealthData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const checkerSheet = ss.getSheetByName(CHECKER_SHEET_NAME);
    if (!checkerSheet || checkerSheet.getLastRow() < 2) {
      return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };
    }

    const data = checkerSheet.getRange(2, 1, checkerSheet.getLastRow() - 1, 11).getValues();

    // Get the date for the most recent Monday.
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize to the start of the day
    const dayOfWeek = today.getDay(); // Sunday = 0, Monday = 1, ...
    const daysSinceMonday = (dayOfWeek + 6) % 7;
    const lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

    let totalChecked = 0;
    let totalDuplicates = 0;

    for (const row of data) {
      const entryDate = new Date(row[0]);
      if (entryDate >= lastMonday) {
        totalChecked++;
        const status = row[10] ? row[10].toString().toLowerCase() : '';
        if (status.includes('duplicate')) {
          totalDuplicates++;
        }
      }
    }

    const percentage = totalChecked > 0 ? Math.round((totalDuplicates / totalChecked) * 100) : 0;
    return { percentage, totalChecked, totalDuplicates };

  } catch (e) {
    return { error: 'Failed to calculate health data: ' + e.message };
  }
}

/**
 * Adds a new record and checks for duplicates using a high-performance QUERY formula.
 * @param {object} formData An object containing firstName, lastName, and dob.
 * @return {object} A result object for the frontend.
 */
function addRecordAndCheckDuplicates(formData) {
  // ** NEW ** Perform a pre-flight check to ensure the Master Sheet is accessible.
  const accessCheck = checkMasterSheetAccess();
  if (!accessCheck.success) {
    const logSheet = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), LOG_SHEET_NAME, ['TIMESTAMP', 'ACTION', 'DETAILS', 'MESSAGE']);
    logEntry(logSheet, 'Access Error', {}, accessCheck.message);
    return { success: false, message: accessCheck.message };
  }

  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const checkerSheet = getOrCreateSheet(activeSs, CHECKER_SHEET_NAME, ['DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'PHONE NUMBER', 'ADDRESS', 'CITY', 'STATE', 'ZIP CODE', 'STATUS']);
    const logSheet = getOrCreateSheet(activeSs, LOG_SHEET_NAME, ['TIMESTAMP', 'ACTION', 'DETAILS', 'MESSAGE']);

    const { firstName, lastName, dob: dobString } = formData; // dobString is in 'YYYY-MM-DD' format

    if (!firstName || !lastName || !dobString) {
      logEntry(logSheet, 'Validation Error', formData, 'Missing mandatory fields.');
      return { success: false, message: 'Error: All fields are required.' };
    }
    
    const duplicatesFound = queryForDuplicates(firstName, lastName, dobString);
    const status = duplicatesFound > 0 ? `${duplicatesFound} Duplicates Found` : 'None';

    const [year, month, day] = dobString.split('-').map(Number);
    const dobForSheet = new Date(year, month - 1, day);

    const newRowData = [ new Date(), '', firstName, lastName, dobForSheet, '', '', '', '', '', status ];
    
    const newRow = checkerSheet.getLastRow() + 1;
    checkerSheet.getRange(newRow, 1, 1, newRowData.length).setValues([newRowData]);

    if (duplicatesFound > 0) {
      checkerSheet.getRange(newRow, 1, 1, newRowData.length).setBackground(NEW_ROW_HIGHLIGHT_COLOR);
      logEntry(logSheet, 'Duplicate Found', formData, status);
      return {
        success: false,
        message: `Record added to "Checker" tab. ${duplicatesFound} duplicates found.`,
        duplicatesFound: duplicatesFound
      };
    } else {
      logEntry(logSheet, 'Record Added', formData, 'No duplicates found.');
      return {
        success: true,
        message: 'Record added to "Checker" tab. No duplicates found.'
      };
    }

  } catch (e) {
    Logger.log('Error in addRecordAndCheckDuplicates: ' + e.stack);
    return { success: false, message: 'An unexpected server error occurred. Please check logs.' };
  }
}


// --- HELPER FUNCTIONS ---

/**
 * ** NEW FUNCTION **
 * Verifies that the Master Spreadsheet and the 'Master' sheet are accessible.
 * @return {{success: boolean, message?: string}}
 */
function checkMasterSheetAccess() {
  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    if (!masterSs) {
      return { success: false, message: 'Error: Master Sheet ID is invalid or you do not have permission.' };
    }
    const masterSheet = masterSs.getSheetByName('Master');
    if (!masterSheet) {
      return { success: false, message: "Error: 'Master' sheet not found in the target spreadsheet." };
    }
    return { success: true };
  } catch (e) {
    Logger.log('Master Sheet access error: ' + e.toString());
    return { success: false, message: 'Error: Cannot access Master Sheet. Check ID and permissions.' };
  }
}

/**
 * Performs a near-instant duplicate check by delegating the search to Google's
 * backend via a temporary QUERY(IMPORTRANGE(...)) formula. This version is
 * designed to exactly mimic the user's proven, working manual query.
 * @param {string} firstName The first name to search for.
 * @param {string} lastName The last name to search for.
 * @param {string} dobString The date of birth to search for as a 'YYYY-MM-DD' string.
 * @return {number} The number of duplicate rows found.
 */
function queryForDuplicates(firstName, lastName, dobString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = getOrCreateSheet(ss, HELPER_SHEET_NAME);
  ss.setActiveSheet(helperSheet);
  helperSheet.hideSheet();

  const safeFirstName = firstName.replace(/'/g, "''");
  const safeLastName = lastName.replace(/'/g, "''");
  
  const selectClause = `SELECT Col3 WHERE LOWER(Col3) = LOWER('${safeFirstName}') AND LOWER(Col4) = LOWER('${safeLastName}') AND Col5 = DATE '${dobString}'`;
  const formula = `=COUNTA(IFERROR(QUERY(IMPORTRANGE("${MASTER_SPREADSHEET_ID}", "Master!A2:G"), "${selectClause}", 0)))`;
  
  const cell = helperSheet.getRange("A1");
  cell.setFormula(formula);
  
  SpreadsheetApp.flush();
  
  try {
    const result = cell.getValue();
    return Number(result) || 0;
  } finally {
    cell.clearContent();
  }
}


/**
 * Gets a sheet by name or creates it with specified headers if it doesn't exist.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {string} sheetName The name of the sheet.
 * @param {Array<string>} headers The headers to add if the sheet is new.
 * @return {Sheet} The sheet object.
 */
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Logs an action to the log sheet.
 * @param {Sheet} logSheet The log sheet object.
 * @param {string} action The action being logged.
 * @param {object} data The form data related to the action.
 * @param {string} message The log message.
 */
function logEntry(logSheet, action, data, message) {
  try {
    const timestamp = new Date();
    const details = JSON.stringify(data);
    logSheet.appendRow([timestamp, action, details, message]);
  } catch(e) {
    Logger.log("Failed to write to log sheet: " + e.message);
  }
}
