




/** ===========================================================================
 * 
 *                        SIDEBAR API FUNCTIONS
 * 
 * ===========================================================================/

/**
 * Determines the context in which the sidebar is running.
 * @return {object} An object describing the context ('MASTER', 'SOURCE', 'OTHER', or 'OTHER_NO_HEADERS').
 */
function getContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssId = ss.getId();
    if (ssId === MASTER_SPREADSHEET_ID) {
      return { context: 'MASTER' };
    }
    if (SOURCE_IDS.indexOf(ssId) !== -1) {
      return { context: 'SOURCE', sourceId: ssId, sourceName: ss.getName() };
    }
    
    // For 'OTHER' context, check for required headers
    var sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getLastColumn() > 0) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
      var hasFirstName = headers.indexOf('FIRST NAME') !== -1;
      var hasLastName = headers.indexOf('LAST NAME') !== -1;
      var hasDob = headers.indexOf('DOB') !== -1;
      
      if (hasFirstName && hasLastName && hasDob) {
        return { context: 'OTHER' };
      }
    }
    
    return { context: 'OTHER_NO_HEADERS' };

  } catch(e) {
    Logger.log('Error in getContext: ' + e.stack);
    return { context: 'ERROR', message: e.message };
  }
}

/**
 * Gathers duplicate health statistics for all source sheets for the admin dashboard.
 * @param {string|null} startDateString The start date in YYYY-MM-DD format, or null.
 * @param {string|null} endDateString The end date in YYYY-MM-DD format, or null.
 * @return {Array<object>} An array of stats objects for each source.
 */
function getAdminDashboardData(startDateString, endDateString) {
  try {
    var sourceData = {};
    SOURCE_IDS.forEach(function(id) {
      try {
        var name = SpreadsheetApp.openById(id).getName();
        sourceData[id] = { sourceId: id, sourceName: name, totalChecked: 0, totalDuplicates: 0 };
      } catch (e) {
        sourceData[id] = { sourceId: id, sourceName: 'Unknown Source (' + id.slice(0, 5) + '...)', totalChecked: 0, totalDuplicates: 0 };
      }
    });

    var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    if (!masterSheet || masterSheet.getLastRow() < 2) return [];

    var masterValues = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
    var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });

    var dateIndex = headers.indexOf('DATE');
    var sourceIdIndex = headers.indexOf('SHEET');
    var potDupIndex = headers.indexOf('POTENTIAL_DUPLICATES');
    if (dateIndex === -1 || sourceIdIndex === -1) return [];

    var useDateFilter = startDateString && endDateString;
    var startDate, endDate;
    if (useDateFilter) {
        startDate = new Date(startDateString);
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date(endDateString);
        endDate.setHours(23, 59, 59, 999);
    }

    masterValues.forEach(function(row) {
      var sourceId = row[sourceIdIndex];
      if (!sourceData[sourceId]) return;

      var entryDate = row[dateIndex] ? new Date(row[dateIndex]) : null;
      if (!entryDate || isNaN(entryDate.getTime())) return;
        
      var isInDateRange = !useDateFilter || (entryDate >= startDate && entryDate <= endDate);

      if (isInDateRange) {
        sourceData[sourceId].totalChecked++;
        if (potDupIndex > -1 && row[potDupIndex] && row[potDupIndex].toString().trim() !== '') {
          sourceData[sourceId].totalDuplicates++;
        }
      }
    });

    return Object.keys(sourceData).map(function(id) {
      var data = sourceData[id];
      return {
        sourceId: data.sourceId,
        sourceName: data.sourceName,
        totalChecked: data.totalChecked,
        totalDuplicates: data.totalDuplicates,
        percentage: data.totalChecked > 0 ? Math.round((data.totalDuplicates / data.totalChecked) * 100) : 0
      };
    });
  } catch (e) {
    Logger.log('Error in getAdminDashboardData: ' + e.stack);
    throw new Error('Could not retrieve dashboard data. ' + e.message);
  }
}


/**
 * Gets duplicate health for a single source sheet.
 * @param {string} sourceId The spreadsheet ID of the source.
 * @return {object} The health data object.
 */
function getSourceSheetHealthData(sourceId) {
    try {
        var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
        var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
        if (!masterSheet || masterSheet.getLastRow() < 2) return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };

        var masterValues = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
        var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
        
        var dateIndex = headers.indexOf('DATE');
        var sourceIdIndex = headers.indexOf('SHEET');
        var potDupIndex = headers.indexOf('POTENTIAL_DUPLICATES');
        if (dateIndex === -1 || sourceIdIndex === -1) return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };
        
        var today = new Date();
        today.setHours(0, 0, 0, 0);
        var dayOfWeek = today.getDay();
        var daysSinceMonday = (dayOfWeek + 6) % 7;
        var lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

        var totalChecked = 0;
        var totalDuplicates = 0;

        masterValues.forEach(function(row) {
            if (row[sourceIdIndex] === sourceId) {
                var entryDate = new Date(row[dateIndex]);
                if (entryDate >= lastMonday) {
                    totalChecked++;
                    if (potDupIndex > -1 && row[potDupIndex] && row[potDupIndex].toString().trim() !== '') {
                        totalDuplicates++;
                    }
                }
            }
        });

        return {
            totalChecked: totalChecked,
            totalDuplicates: totalDuplicates,
            percentage: totalChecked > 0 ? Math.round((totalDuplicates / totalChecked) * 100) : 0
        };
    } catch (e) {
        Logger.log('Error in getSourceSheetHealthData: ' + e.stack);
        throw new Error('Could not get source sheet health data. ' + e.message);
    }
}


/**
 * Adds a new record to the currently active source sheet.
 * @param {object} formData An object containing firstName, lastName, and dob.
 * @return {object} A result object for the frontend.
 */
function addRecordToSourceSheet(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
    
    var firstNameCol = headers.indexOf('FIRST NAME') + 1;
    var lastNameCol = headers.indexOf('LAST NAME') + 1;
    var dobCol = headers.indexOf('DOB') + 1;

    if (!firstNameCol || !lastNameCol || !dobCol) {
      return { success: false, message: 'Could not find First Name, Last Name, and DOB columns.' };
    }
    
    var newRow = sheet.getLastRow() + 1;
    var [year, month, day] = formData.dob.split('-').map(Number);
    var dobDate = new Date(year, month - 1, day);

    sheet.getRange(newRow, firstNameCol).setValue(formData.firstName);
    sheet.getRange(newRow, lastNameCol).setValue(formData.lastName);
    sheet.getRange(newRow, dobCol).setValue(dobDate).setNumberFormat('MM/dd/yyyy');

    // Set the date field if it exists
    var dateCol = headers.indexOf('DATE') + 1;
    if (dateCol) {
      sheet.getRange(newRow, dateCol).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
    }

    return { success: true, message: 'Record added to this sheet. Syncing...' };
  } catch (e) {
    Logger.log('Error in addRecordToSourceSheet: ' + e.stack);
    return { success: false, message: 'An unexpected server error occurred.' };
  }
}


/**
 * [CONTEXT: OTHER] Calculates and returns duplicate health statistics for the sidebar.
 * This is now updated to query the Master sheet for overall system health.
 * @return {object} An object containing percentage, totalChecked, and totalDuplicates.
 */
function getDuplicateHealthData() {
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    if (!masterSheet || masterSheet.getLastRow() < 2) {
      return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };
    }

    var masterValues = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
    var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
    
    var dateIndex = headers.indexOf('DATE');
    var potDupIndex = headers.indexOf('POTENTIAL_DUPLICATES');
    if (dateIndex === -1) return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var dayOfWeek = today.getDay();
    var daysSinceMonday = (dayOfWeek + 6) % 7;
    var lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

    var totalChecked = 0;
    var totalDuplicates = 0;

    masterValues.forEach(function(row) {
      var entryDate = new Date(row[dateIndex]);
      if (entryDate >= lastMonday) {
        totalChecked++;
        if (potDupIndex > -1 && row[potDupIndex] && row[potDupIndex].toString().trim() !== '') {
          totalDuplicates++;
        }
      }
    });

    return {
      totalChecked: totalChecked,
      totalDuplicates: totalDuplicates,
      percentage: totalChecked > 0 ? Math.round((totalDuplicates / totalChecked) * 100) : 0
    };
  } catch (e) {
    Logger.log('Error in getDuplicateHealthData (querying master): ' + e.stack);
    throw new Error('Failed to calculate health data. ' + e.message);
  }
}

/**
 * [CONTEXT: OTHER] Adds/updates a record, creating a log in both the local and master sheet.
 * If a duplicate is found (First/Last/DOB match), it updates the existing Master record.
 * If not found, it creates a new record in the Master sheet.
 * It ALWAYS adds a log entry to the 'Checker' sheet in both the local and master spreadsheets.
 *
 * @param {object} formData An object containing firstName, lastName, and dob.
 * @return {object} A result object for the frontend.
 */
function addRecordAndCheckDuplicates(formData) {
  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    
    const checkerHeaders = ['DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'PHONE NUMBER', 'ADDRESS', 'CITY', 'STATE', 'ZIP CODE', 'STATUS'];
    const masterCheckerHeaders = checkerHeaders.concat(['SOURCE_SHEET_NAME']);
    
    // Get/Create BOTH local and master checker sheets
    const localCheckerSheet = getOrCreateSheet(activeSs, CHECKER_SHEET_NAME, checkerHeaders);
    const masterCheckerSheet = getOrCreateSheet(masterSs, CHECKER_SHEET_NAME, masterCheckerHeaders);

    const { firstName, lastName, dob: dobString } = formData;

    if (!firstName || !lastName || !dobString) {
      return { success: false, message: 'Error: All fields are required.' };
    }
    
    const masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    if (!masterSheet) throw new Error('Master sheet not found.');

    const masterHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    const dateColIndex1 = masterHeaders.indexOf('DATE') + 1;
    const dobColIndex1 = masterHeaders.indexOf('DOB') + 1;

    const existingRecord = findMasterRecordByDetails_(firstName, lastName, dobString);
    const [year, month, day] = dobString.split('-').map(Number);
    const dobForSheet = new Date(year, month - 1, day);
    
    let status = '';
    let duplicatesFound = 0;

    if (existingRecord && existingRecord.uuid) {
      // DUPLICATE FOUND: Update the existing record in the Master sheet
      const masterRow = existingRecord.row;
      const rangeToUpdate = masterSheet.getRange(masterRow, 1, 1, masterSheet.getLastColumn());
      rangeToUpdate.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      masterSheet.getRange(masterRow, dateColIndex1).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
      status = `Existing Master record updated (Row ${masterRow})`;
      duplicatesFound = 1;
    } else {
      // NEW RECORD: Append a new row to the Master sheet
      const newUuid = Utilities.getUuid();
      const newMasterRowData = [
        new Date(), '', firstName, lastName, dobForSheet, '', 
        'manual_sidebar_entry', '', 'MANUAL', activeSs.getName(), newUuid, ''
      ];
      masterSheet.appendRow(newMasterRowData);
      const newMasterRow = masterSheet.getLastRow();
      const newRowRange = masterSheet.getRange(newMasterRow, 1, 1, masterSheet.getLastColumn());
      newRowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      if (dateColIndex1 > 0) newRowRange.getCell(1, dateColIndex1).setNumberFormat('MM/dd/yyyy');
      if (dobColIndex1 > 0) newRowRange.getCell(1, dobColIndex1).setNumberFormat('MM/dd/yyyy');

      status = 'New record added to Master';
    }

    // ALWAYS add a row to the LOCAL Checker sheet for logging purposes
    const newCheckerRowData = [new Date(), '', firstName, lastName, dobForSheet, '', '', '', '', '', status];
    const newLocalCheckerRow = localCheckerSheet.getLastRow() + 1;
    const localCheckerRange = localCheckerSheet.getRange(newLocalCheckerRow, 1, 1, newCheckerRowData.length)
    localCheckerRange.setValues([newCheckerRowData]);
    localCheckerRange.getCell(1,1).setNumberFormat('MM/dd/yyyy');
    localCheckerRange.getCell(1,5).setNumberFormat('MM/dd/yyyy');

    // ALWAYS add a row to the MASTER Checker sheet for logging purposes
    const newMasterCheckerRowData = [new Date(), '', firstName, lastName, dobForSheet, '', '', '', '', '', status, activeSs.getName()];
    const newMasterCheckerRow = masterCheckerSheet.getLastRow() + 1;
    const masterCheckerRange = masterCheckerSheet.getRange(newMasterCheckerRow, 1, 1, newMasterCheckerRowData.length)
    masterCheckerRange.setValues([newMasterCheckerRowData]);
    masterCheckerRange.getCell(1,1).setNumberFormat('MM/dd/yyyy');
    masterCheckerRange.getCell(1,5).setNumberFormat('MM/dd/yyyy');

    if (duplicatesFound > 0) {
      localCheckerRange.setBackground(NEW_ROW_DATE_HIGHLIGHT);
      masterCheckerRange.setBackground(NEW_ROW_DATE_HIGHLIGHT);
      return {
        success: false, // Return 'false' to trigger the warning color on the frontend
        message: status,
        duplicatesFound: duplicatesFound
      };
    } else {
      return {
        success: true,
        message: status
      };
    }

  } catch (e) {
    Logger.log('Error in addRecordAndCheckDuplicates: ' + e.stack);
    return { success: false, message: 'An unexpected server error occurred. Please check logs.' };
  }
}

/**
 * Uses the Gemini API to generate a natural language summary of the admin dashboard stats.
 * @param {Array<object>} stats The array of stats objects from getAdminDashboardData.
 * @return {string} A text summary generated by the AI.
 */
function getGeminiHealthSummary(stats) {
  if (!stats || stats.length === 0) {
    return 'No data available to analyze.';
  }
  try {
    var apiKey = getApiKey_(); // Will throw an error if not set
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

    var prompt = "You are a data quality analyst for a company called BHG. " +
                 "Based on the following JSON data of duplicate record percentages from various data sources, " +
                 "provide a concise, professional summary (2-3 sentences) for an executive. " +
                 "Highlight the source with the highest duplicate rate (the biggest problem) and the one with the lowest (the healthiest). " +
                 "Do not use markdown or formatting. Be direct and start the summary immediately.\n\n" +
                 "Data: " + JSON.stringify(stats);

    var payload = {
      "contents": [{
        "parts": [{ "text": prompt }]
      }]
    };

    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      var jsonResponse = JSON.parse(responseBody);
      // As per Gemini REST API docs
      var text = jsonResponse.candidates[0].content.parts[0].text;
      return text.trim();
    } else {
      Logger.log('Gemini API Error: ' + responseBody);
      var errorJson = JSON.parse(responseBody);
      var errorMessage = errorJson.error ? errorJson.error.message : 'An unknown API error occurred.';
      throw new Error('Gemini API Error: ' + errorMessage);
    }
  } catch (e) {
    Logger.log('Error in getGeminiHealthSummary: ' + e.stack);
    // Rethrow a user-friendly message
    throw new Error('Could not get AI summary. ' + e.message);
  }
}

/**
 * A debug function callable from source sheets to verify library connection and master sheet access.
 * This is crucial for troubleshooting "Could not connect" errors.
 * @return {object} A result object with debug information.
 */
function debugConnection() {
  try {
    const user = Session.getActiveUser().getEmail();
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetId = activeSs.getId();
    const activeSheetName = activeSs.getName();
    let masterAccessStatus = 'Unknown';
    let masterAccessMessage = '';

    try {
      const masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
      const masterName = masterSs.getName();
      masterAccessStatus = 'Success';
      masterAccessMessage = `Successfully accessed Master Sheet: "${masterName}".`;
    } catch (e) {
      masterAccessStatus = 'FAILED';
      masterAccessMessage = `Could not open the Master Sheet. The user ${user} may not have at least VIEW permissions. Error: ${e.message}`;
    }
    
    return {
      status: 'OK',
      user: user,
      masterAccess: masterAccessStatus,
      activeSheetId: activeSheetId,
      activeSheetName: activeSheetName,
      message: 'Library function executed successfully. ' + masterAccessMessage
    };

  } catch (e) {
    return {
      status: 'ERROR',
      user: Session.getActiveUser() ? Session.getActiveUser().getEmail() : 'Unknown',
      masterAccess: 'Not Tested',
      activeSheetId: SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : 'Unknown',
      activeSheetName: SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getName() : 'Unknown',
      message: 'A critical error occurred within the debug function itself: ' + e.message
    };
  }
}