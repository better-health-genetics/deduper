/****************************************************************************************************
 * ___  ____ ___ ___ ____ ____    _  _ ____ ____ _    ___ _  _    ____ ____ _  _ ____ ___ _ ____ ____ 
 * |__] |___  |   |  |___ |__/    |__| |___ |__| |     |  |__|    | __ |___ |\ | |___  |  | |    [__  
 * |__] |___  |   |  |___ |  \    |  | |___ |  | |___  |  |  |    |__] |___ | \| |___  |  | |___ ___] 
 * 
 *                             BETTER HEALTH GENETICS — DATA CONSOLIDATOR
 ****************************************************************************************************/

// =================================================================================================
// MERGED SCRIPT: This file now contains both the original Data Consolidator backend logic
// AND the backend functions required to power the "BHG DeDuper" sidebar UI.
// =================================================================================================


/* ---------- CONFIG ---------- */
var LIVE_SOURCE_IDS = [
  '1URF6p99IPY8_kPCNOhYigL_S7UMvyQAdCyB1dCtVGrY', //MDS Unity (New)
  '1bgOimFOdjXcFZqWzWVvnUw1xRnyhWCAdGrxy1JSlVR4', //YEST
  '1eXJzHNsitv86igh8oAg-HAmz3LXqT9A-aCz9pl3zEvo', //RCC
  '1YscvMF-vB7HgfzYrVCWDCIxGAuerWyfcG2_jZgDttVA', //Wave
  '1NzoxehLtg55gXgBEyZTwGxdl2wltioGnv9Dn_7-HwhI', //TBM - ADL
  '10WVTn2rtFjLJzN4943Nk4G8JsOm2qt9Sa1S6-aDU6ZM', //MDS ADL (New)
  '1hKtLqUgaGhHZvcFm_BydLBYAPUJ9IcyYtDhZGezy2u8', //MDS EVEREST
  '1D207Z49WrW7UgEvbvCRvqB_uVP1IEu1lB89Vxkqt1gY', //MDS Pathway (New)
  '19PfI2wDAUB_qKFHkMJzUHIkLSwzaNmS91h63P1aHrvw', //MDS Star (New)
  '1jlSPG05hqVh-RBsDzFmdYT_e2eA2cMEz9A4quaeKG8I', //TBM EVEREST
  '1XDZidmShTfyku6rWH5YYJBWYTdAL3RlnqZyByqGYRXo', //TBM District
  '1jodXK36Y7ojU7g9iij6FAZCqLS37BTY407Hgv00SWAk', //TBM PATH (New)
  '1ATCASWpQ3Ju5481B-Uh_1fhIDfAafnxnKd1903K5jA0', //TBM - STARLABS
  '1ZX36vmCYR0D5_8Y-fx9sLct46As_bMD6mGIxm9aBPN8'  //TBM UNITY
];

var SOURCE_IDS = [
  '1QgLFeJiw8vuF849rGokWEDDsKyIpd5iKY9OY_SdrA4Q', //DEV RCC
  '14NZWsm3HVlEDVwV-ncWnZRF3GK3hKByen0euSbd6CUg', //DEV TBM
];
var LIVE_SPREADSHEET_ID = '161uw5s1lOwhV7YTX8uLKDw6TMl_QLEpSqxmHEftZzVA';
var MASTER_SPREADSHEET_ID = '14I2UGjK3Vmsbya9PJZY16SxIDKotply3ysLHkWI-Hpw';

var DEST_SHEET_NAME = 'Master';
var LOG_SHEET_NAME = 'Log';
var DUPLICATES_SHEET_NAME = 'Duplicates';
var CHECKER_SHEET_NAME = 'Checker';
var HELPER_SHEET_NAME = '_QueryHelper';

var MANDATORY_FIELDS = ['DATE', 'FIRST NAME', 'LAST NAME', 'DOB'];
var NEW_ROW_DATE_HIGHLIGHT = '#ffcccc';

var PROP_PREFIX = 'lastRow_';
var CONS_PROP_PREFIX = 'consolidate_lastRow_';
var DUP_GROUP_PREFIX = 'dup_logged_';

var LOG_HEADERS = ['TIMESTAMP', 'REASON', 'SOURCE_SPREADSHEET_ID', 'SOURCE_SHEET_NAME', 'ROW_NUMBER', 'ORIGINAL_VALUES', 'LINK_TO_ROW', 'DUPLICATE_LINKS', 'TRIGGERING_USER'];
/* ------------------------------- */

/** ===========================================================================
 * 
 *                        MERGED UI & MENU FUNCTIONS
 * 
 * ===========================================================================/

/**
 * Auto-add menu and open sidebar when the Master spreadsheet opens.
 * This function merges the logic from both the consolidator and the sidebar.
 */
function onOpen() {
  // Add the main "Data Consolidator" menu for backend tasks.
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Data Consolidator')
      .addItem('🔄 Rebuild Master (Full Reset)', 'ui_rebuildMasterFullReset_')
      .addItem('📥 Reimport History (Keep Master)', 'ui_reimportHistoryNoClear_')
      .addSeparator()
      .addItem('Show DeDuper Sidebar', 'showDuplicateCheckerSidebar') // Added sidebar opener
      .addSeparator()
      .addItem('🧭 Seed Source Last Rows', 'ui_initSourceLastRows_')
      .addItem('🧭 Seed Consolidation Cursors', 'ui_initConsolidateLastRows_')
      .addSeparator()
      .addItem('⚙️ Create Source onEdit Triggers', 'ui_createTriggersForSources_')
      .addItem('⏰ Create Daily Consolidation (2am)', 'ui_createDailyConsolidationTrigger_')
      .addSeparator()
      .addItem('📜 Open Log Sheet', 'ui_openLogSheet_')
      .addToUi();
  } catch (e) {
    Logger.log('addConsolidatorMenu error: ' + e.message);
  }

  // Automatically open the sidebar for daily use.
  showDuplicateCheckerSidebar();
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


/** ===========================================================================
 * 
 *                        SIDEBAR API FUNCTIONS
 * 
 * ===========================================================================/

/**
 * Determines the context in which the sidebar is running.
 * @return {object} An object describing the context ('MASTER', 'SOURCE', or 'OTHER').
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
    return { context: 'OTHER' };
  } catch(e) {
    return { context: 'ERROR', message: e.message };
  }
}

/**
 * Gathers duplicate health statistics for all source sheets for the admin dashboard.
 * @return {Array<object>} An array of stats objects for each source.
 */
function getAdminDashboardData() {
  try {
    var sourceData = {};
    SOURCE_IDS.forEach(function(id) {
        try {
            var name = SpreadsheetApp.openById(id).getName();
            sourceData[id] = { sourceName: name, totalChecked: 0, totalDuplicates: 0 };
        } catch(e) {
             sourceData[id] = { sourceName: 'Unknown Source ('+id.slice(0,5)+'...)', totalChecked: 0, totalDuplicates: 0 };
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

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var dayOfWeek = today.getDay();
    var daysSinceMonday = (dayOfWeek + 6) % 7;
    var lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

    masterValues.forEach(function(row) {
      var entryDate = new Date(row[dateIndex]);
      var sourceId = row[sourceIdIndex];
      if (entryDate >= lastMonday && sourceData[sourceId]) {
        sourceData[sourceId].totalChecked++;
        if (row[potDupIndex] && row[potDupIndex].toString().trim() !== '') {
          sourceData[sourceId].totalDuplicates++;
        }
      }
    });

    return Object.keys(sourceData).map(function(id) {
      var data = sourceData[id];
      return {
        sourceName: data.sourceName,
        totalChecked: data.totalChecked,
        totalDuplicates: data.totalDuplicates,
        percentage: data.totalChecked > 0 ? Math.round((data.totalDuplicates / data.totalChecked) * 100) : 0
      };
    });
  } catch (e) {
    Logger.log('Error in getAdminDashboardData: ' + e.stack);
    return [];
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
                    if (row[potDupIndex] && row[potDupIndex].toString().trim() !== '') {
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
        return { percentage: 0, totalChecked: 0, totalDuplicates: 0 };
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
    sheet.getRange(newRow, dobCol).setValue(dobDate);

    // Set the date field if it exists
    var dateCol = headers.indexOf('DATE') + 1;
    if (dateCol) {
      sheet.getRange(newRow, dateCol).setValue(new Date());
    }

    return { success: true, message: 'Record added to this sheet. Syncing...' };
  } catch (e) {
    Logger.log('Error in addRecordToSourceSheet: ' + e.stack);
    return { success: false, message: 'An unexpected server error occurred.' };
  }
}


/**
 * [CONTEXT: OTHER] Calculates and returns duplicate health statistics for the sidebar.
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

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dayOfWeek = today.getDay();
    const daysSinceMonday = (dayOfWeek + 6) % 7;
    const lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

    let totalChecked = 0;
    let totalDuplicates = 0;

    for (const row of data) {
      const entryDate = new Date(row[0]);
      if (entryDate >= lastMonday) {
        totalChecked++;
        const status = row[10] ? row[10].toString().toLowerCase() : '';
        if (status.includes('duplicate') || status.includes('updated')) {
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
 * [CONTEXT: OTHER] Adds or updates a record from the sidebar, using the MASTER_UUID system.
 * If a duplicate is found (First/Last/DOB match), it updates the existing Master record.
 * If not found, it creates a new record in the Master sheet.
 * It ALWAYS adds a log entry to the 'Checker' sheet.
 *
 * @param {object} formData An object containing firstName, lastName, and dob.
 * @return {object} A result object for the frontend.
 */
function addRecordAndCheckDuplicates(formData) {
  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const checkerSheet = getOrCreateSheet(activeSs, CHECKER_SHEET_NAME, ['DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'PHONE NUMBER', 'ADDRESS', 'CITY', 'STATE', 'ZIP CODE', 'STATUS']);
    
    const { firstName, lastName, dob: dobString } = formData;

    if (!firstName || !lastName || !dobString) {
      return { success: false, message: 'Error: All fields are required.' };
    }
    
    const masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    if (!masterSheet) throw new Error('Master sheet not found.');

    const existingRecord = findMasterRecordByDetails_(firstName, lastName, dobString);
    const [year, month, day] = dobString.split('-').map(Number);
    const dobForSheet = new Date(year, month - 1, day);
    
    let status = '';
    let duplicatesFound = 0;

    if (existingRecord && existingRecord.uuid) {
      // DUPLICATE FOUND: Update the existing record in the Master sheet
      const masterRow = existingRecord.row;
      masterSheet.getRange(masterRow, 1).setValue(new Date()); // Update the main DATE field
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
      status = 'New record added to Master';
    }

    // ALWAYS add a row to the Checker sheet for logging purposes
    const newCheckerRowData = [new Date(), '', firstName, lastName, dobForSheet, '', '', '', '', '', status];
    const newCheckerRow = checkerSheet.getLastRow() + 1;
    checkerSheet.getRange(newCheckerRow, 1, 1, newCheckerRowData.length).setValues([newCheckerRowData]);

    if (duplicatesFound > 0) {
      checkerSheet.getRange(newCheckerRow, 1, 1, newCheckerRowData.length).setBackground(NEW_ROW_DATE_HIGHLIGHT);
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
 * Helper for the sidebar: Finds a record in the Master sheet by First, Last, and DOB.
 * Uses a fast query to find the MASTER_UUID, then finds the row for that UUID.
 * @return {{uuid: string, row: number} | null} The record's UUID and row number, or null if not found.
 */
function findMasterRecordByDetails_(firstName, lastName, dobString) {
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const helperSheet = getOrCreateSheet(ss, HELPER_SHEET_NAME);
  helperSheet.hideSheet();

  const safeFirstName = firstName.replace(/'/g, "''");
  const safeLastName = lastName.replace(/'/g, "''");

  const selectClause = `SELECT K WHERE LOWER(C) = LOWER('${safeFirstName}') AND LOWER(D) = LOWER('${safeLastName}') AND E = DATE '${dobString}' LIMIT 1`;
  const formula = `=IFERROR(QUERY(Master!A2:L, "${selectClause}", 0), "")`;
  
  const cell = helperSheet.getRange("A1");
  try {
    cell.setFormula(formula);
    SpreadsheetApp.flush();
    const uuid = cell.getValue().toString().trim();
    
    if (!uuid) {
      return null;
    }

    const masterSheet = ss.getSheetByName(DEST_SHEET_NAME);
    const uuidColValues = masterSheet.getRange(2, 11, masterSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < uuidColValues.length; i++) {
      if (uuidColValues[i][0] === uuid) {
        return { uuid: uuid, row: i + 2 };
      }
    }
    return null;
  } finally {
    cell.clearContent();
  }
}

/** ===========================================================================
 * 
 *                  ORIGINAL DATA CONSOLIDATOR SCRIPT
 * 
 * ===========================================================================/

/* ----------------- Helpers (single source of truth for headers/indices) ----------------- */
function applyOverflowWrapStrategy_(sheet, headerName) {
  if (!sheet || !headerName) return;
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
    var colIndex = headers.indexOf(headerName.toUpperCase());

    if (colIndex !== -1) {
      var col = colIndex + 1;
      var numRows = Math.max(1, sheet.getMaxRows() - 1);
      var range = sheet.getRange(2, col, numRows, 1);
      if (SpreadsheetApp.WrapStrategy && SpreadsheetApp.WrapStrategy.OVERFLOW) {
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      } else {
        range.setWrap(true); 
      }
    }
  } catch (e) {
    Logger.log('Failed to apply wrap strategy for header "' + headerName + '" on sheet "' + sheet.getName() + '": ' + e.message);
  }
}


function clearSheetKeepHeader_(sheet, minCols) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(minCols || sheet.getLastColumn(), 1);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent().clearNote().setBackground(null);
  }
}

function clearColumnByHeader_(sheet, headerName) {
  if (!sheet) return;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return (h||'').toString().trim().toUpperCase(); });
  var idx = headers.indexOf((headerName||'').toString().trim().toUpperCase());
  if (idx !== -1) {
    var lastRow = Math.max(2, sheet.getLastRow());
    sheet.getRange(2, idx + 1, lastRow - 1, 1).clearContent().clearNote();
  }
}

function resetConsolidationCursors_() {
  var props = PropertiesService.getScriptProperties();
  SOURCE_IDS.forEach(function(spreadId) {
    try {
      var ss = SpreadsheetApp.openById(spreadId);
      ss.getSheets().forEach(function(sh) {
        var k = CONS_PROP_PREFIX + ss.getId() + '_' + sh.getSheetId();
        props.setProperty(k, '1');
        Logger.log('Reset ' + k + ' => 1');
      });
    } catch (e) {
      Logger.log('resetConsolidationCursors_ error for ' + spreadId + ': ' + e.message);
    }
  });
}

function resetDuplicateGroupLog_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  Object.keys(all).forEach(function(k) {
    if (k && k.indexOf(DUP_GROUP_PREFIX) === 0) {
      props.deleteProperty(k);
      Logger.log('Deleted dup-log key: ' + k);
    }
  });
}

function resetMasterAndReimportAll() {
  var destSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var expectedHeaders = [
    'DATE','TEST TYPE','FIRST NAME','LAST NAME','DOB','ZIP CODE',
    'SHEET','SHEET LINK TO ROW','LAB (SHEET NAME)','GROUP (FILE NAME)','MASTER_UUID','POTENTIAL_DUPLICATES'
  ];

  var master = destSs.getSheetByName(DEST_SHEET_NAME) || destSs.insertSheet(DEST_SHEET_NAME);
  var log    = destSs.getSheetByName(LOG_SHEET_NAME) || destSs.insertSheet(LOG_SHEET_NAME);
  var dups   = destSs.getSheetByName(DUPLICATES_SHEET_NAME) || destSs.insertSheet(DUPLICATES_SHEET_NAME);

  try {
    var f = master.getFilter && master.getFilter();
    if (f) f.remove();
    master.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    var curCols = master.getLastColumn();
    if (curCols > expectedHeaders.length) {
      master.deleteColumns(expectedHeaders.length + 1, curCols - expectedHeaders.length);
    } else if (curCols < expectedHeaders.length) {
      master.insertColumnsAfter(curCols, expectedHeaders.length - curCols);
    }
    var maxRows = Math.max(2, master.getMaxRows());
    var numRows = maxRows - 1;
    if (numRows > 0) {
      master.getRange(2, 1, numRows, expectedHeaders.length)
            .clearContent()
            .clearNote()
            .setBackground(null);
    }
  } catch (e) {
    Logger.log('Master cleanup failed (pre): ' + e.message);
  }

  applyOverflowWrapStrategy_(master, 'POTENTIAL_DUPLICATES');
  applyOverflowWrapStrategy_(master, 'MASTER_UUID');


  ensureDestinationHeader(log, ['TIMESTAMP','REASON','SOURCE_SPREADSHEET_ID','SOURCE_SHEET_NAME','ROW_NUMBER','ORIGINAL_VALUES','LINK_TO_ROW','DUPLICATE_LINKS','TRIGGERING_USER']);
  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, Math.max(1, log.getLastColumn()))
       .clearContent()
       .clearNote();
  }

  var dupHeaders = ['FIRST NAME','LAST NAME','DOB','COUNT','LINKS'];
  ensureDestinationHeader(dups, dupHeaders);
  if (dups.getLastRow() > 1) {
    dups.getRange(2, 1, dups.getLastRow() - 1, Math.max(dups.getLastColumn(), dupHeaders.length))
        .clearContent()
        .clearNote();
  }

  resetConsolidationCursors_();
  resetDuplicateGroupLog_();
  Logger.log('After reset, Master lastRow=' + master.getLastRow() + ', lastCol=' + master.getLastColumn());
  consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME);

  try {
    var tz = destSs.getSpreadsheetTimeZone();
    findAndMarkDuplicates(master, expectedHeaders, tz);
  } catch (e) {
    Logger.log('Post-rebuild duplicate pass failed: ' + e.message);
  }
  Logger.log('resetMasterAndReimportAll complete.');
}

function shouldDebounceRow_(spreadId, sheetId, rowNum, windowMs) {
  try {
    var key = 'debounce_' + spreadId + '_' + sheetId + '_' + rowNum;
    var cache = CacheService.getScriptCache();
    var hit = cache.get(key);
    if (hit) return true;
    cache.put(key, '1', Math.max(1, Math.floor(windowMs/1000)));
    return false;
  } catch (e) { return false; }
}

function isRelevantColumn_(headerUpperRow, colNumber1Based) {
  var idx = colNumber1Based - 1;
  var name = (headerUpperRow[idx] || '').toString().trim().toUpperCase().replace(/\s+/g,' ');
  return ['DATE','TEST TYPE','FIRST NAME','LAST NAME','DOB','ZIP CODE','MASTER_UUID','POTENTIAL_DUPLICATES'].indexOf(name) !== -1;
}

function resetCursorsAndReimportWithoutClearing() {
  resetConsolidationCursors_();
  resetDuplicateGroupLog_();
  consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME);
  Logger.log('resetCursorsAndReimportWithoutClearing complete.');
}

function ensureHeaderAndReturnIndex(sheet, headerName) {
  try {
    var lastCol = Math.max(1, sheet.getLastColumn());
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return (h||'').toString().trim(); });
    for (var i = 0; i < headerRow.length; i++) {
      if ((headerRow[i] || '').toString().trim().toUpperCase() === headerName.toString().trim().toUpperCase()) return i;
    }
    var target = headerName.toString().trim().toUpperCase().replace(/\s+/g,'');
    for (var j = 0; j < headerRow.length; j++) {
      if (((headerRow[j]||'').toString().trim().toUpperCase().replace(/\s+/g,'')) === target) return j;
    }
    var writeCol = lastCol + 1;
    sheet.getRange(1, writeCol).setValue(headerName);
    return writeCol - 1;
  } catch (e) {
    Logger.log('ensureHeaderAndReturnIndex error: ' + e.message);
    return -1;
  }
}

function indexOfHeader(headerArray, headerName) {
  var up = headerName.toString().trim().toUpperCase();
  var idx = headerArray.indexOf(up);
  if (idx !== -1) return idx;
  var target = up.replace(/\s+/g, '');
  for (var i = 0; i < headerArray.length; i++) {
    var cand = (headerArray[i] || '').toString().trim().toUpperCase().replace(/\s+/g, '');
    if (cand === target) return i;
  }
  return -1;
}

function normalizeDateValue(value, defaultYear) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var s = value.toString().trim();
  var today = new Date();
  defaultYear = defaultYear || today.getFullYear();

  var m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    var month = parseInt(m[1], 10);
    var day = parseInt(m[2], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      var candidate = new Date(defaultYear, month - 1, day);
      if (candidate.getTime() > today.getTime()) candidate = new Date(defaultYear - 1, month - 1, day);
      return candidate;
    }
  }

  var m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m2) {
    var mm = parseInt(m2[1], 10);
    var dd = parseInt(m2[2], 10);
    var yy = parseInt(m2[3], 10);
    if (yy < 100) yy += (yy >= 70 ? 1900 : 2000);
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) return new Date(yy, mm - 1, dd);
  }

  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return parsed;
  return s;
}

function findMasterRowByUuid(masterSheet, uuid, uuidColIndex) {
  if (!masterSheet || !uuid) return -1;
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return -1;

  var count = lastRow - 1;
  var vals = masterSheet.getRange(2, uuidColIndex, count, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if ((vals[i][0] || '').toString().trim() === uuid) return 2 + i;
  }
  return -1;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function initSourceLastRows() {
  var props = PropertiesService.getScriptProperties();
  SOURCE_IDS.forEach(function(spreadId) {
    try {
      var ss = SpreadsheetApp.openById(spreadId);
      var sheets = ss.getSheets();
      sheets.forEach(function(sh) {
        var key = PROP_PREFIX + ss.getId() + '_' + sh.getSheetId();
        var lr = sh.getLastRow();
        props.setProperty(key, String(lr));
        Logger.log('Set prop ' + key + ' => ' + lr);
      });
    } catch (e) {
      Logger.log('initSourceLastRows error for ' + spreadId + ': ' + e.message);
    }
  });
  Logger.log('initSourceLastRows finished.');
}

function initConsolidateLastRows() {
  var props = PropertiesService.getScriptProperties();
  SOURCE_IDS.forEach(function(spreadId) {
    try {
      var ss = SpreadsheetApp.openById(spreadId);
      var sheets = ss.getSheets();
      sheets.forEach(function(sh) {
        var key = CONS_PROP_PREFIX + ss.getId() + '_' + sh.getSheetId();
        var lr = sh.getLastRow();
        if (lr < 1) lr = 1;
        props.setProperty(key, String(lr));
        Logger.log('Set consolidate prop ' + key + ' => ' + lr);
      });
    } catch (e) {
      Logger.log('initConsolidateLastRows error for ' + spreadId + ': ' + e.message);
    }
  });
  Logger.log('initConsolidateLastRows finished.');
}

function createTriggersForSources() {
  SOURCE_IDS.forEach(function(spreadId) {
    try {
      var ss = SpreadsheetApp.openById(spreadId);
      var sourceId = ss.getId();

      var all = ScriptApp.getProjectTriggers();
      all.forEach(function(t) {
        try {
          if (t.getHandlerFunction() === 'onEditSourceInstallable') {
            var tid = t.getTriggerSourceId ? t.getTriggerSourceId() : null;
            if (tid === sourceId) ScriptApp.deleteTrigger(t);
          }
        } catch (e) {}
      });

      ScriptApp.newTrigger('onEditSourceInstallable')
        .forSpreadsheet(ss)
        .onEdit()
        .create();

      Logger.log('Created trigger for: ' + ss.getName() + ' (' + sourceId + ')');
    } catch (err) {
      Logger.log('Failed to create trigger for ' + spreadId + ': ' + err.message);
    }
  });
  Logger.log('createTriggersForSources finished.');
}

function createDailyConsolidationTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t) {
    try {
      if (t.getHandlerFunction() === 'runForProvidedSheetsIncremental') ScriptApp.deleteTrigger(t);
    } catch (e) {}
  });
  ScriptApp.newTrigger('runForProvidedSheetsIncremental')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('Daily consolidation trigger created (runForProvidedSheetsIncremental).');
}

/**
 * ** REFACTORED onEdit TRIGGER **
 * This function now prioritizes MASTER_UUID for updates. If a UUID is present on
 * an edited row, it updates the corresponding Master record. If no UUID is present,
 * it ALWAYS creates a new record in Master, preventing accidental merges.
 */
function onEditSourceInstallable(e) {
  var lock;
  try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000)) { // Increased lock wait time
        Logger.log('onEditSourceInstallable could not get lock for ' + e.source.getId());
        return;
    }
    
    var props = PropertiesService.getScriptProperties();
    var ss = e.source;
    var sheet = e.range.getSheet();
    var spreadId = ss.getId();
    var sheetId = sheet.getSheetId();
    var editedRow = e.range.getRow();
    var editedCol = e.range.getColumn();

    if (editedRow === 1) return;

    var numCols = Math.max(1, sheet.getLastColumn());
    var headerUpper = sheet.getRange(1, 1, 1, numCols).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });

    if (!isRelevantColumn_(headerUpper, editedCol)) return;
    if (shouldDebounceRow_(spreadId, sheetId, editedRow, 3000)) return;

    var propKey = PROP_PREFIX + spreadId + '_' + sheetId;
    var prevLastRow = Number(props.getProperty(propKey)) || 0;
    var currentLastRow = sheet.getLastRow();

    if (!props.getProperty(propKey)) {
        props.setProperty(propKey, String(currentLastRow));
        return;
    }
    if (currentLastRow < prevLastRow) {
        props.setProperty(propKey, String(currentLastRow));
    }

    var header = headerUpper; // Use already fetched header
    var idxFirst = indexOfHeader(header, 'FIRST NAME');
    var idxLast = indexOfHeader(header, 'LAST NAME');
    var idxDOB = indexOfHeader(header, 'DOB');
    var idxDate = indexOfHeader(header, 'DATE');
    var idxZip = indexOfHeader(header, 'ZIP CODE');
    var idxTest = indexOfHeader(header, 'TEST TYPE');

    if (idxFirst === -1 || idxLast === -1 || idxDOB === -1) {
        props.setProperty(propKey, String(currentLastRow));
        return;
    }

    var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    var masterTz = masterSs.getSpreadsheetTimeZone();
    var masterMap = buildMasterNameDobMap(MASTER_SPREADSHEET_ID); // For duplicate suggestions

    var destHeaders = ['DATE','TEST TYPE','FIRST NAME','LAST NAME','DOB','ZIP CODE','SHEET','SHEET LINK TO ROW','LAB (SHEET NAME)','GROUP (FILE NAME)','MASTER_UUID','POTENTIAL_DUPLICATES'];
    ensureDestinationHeader(masterSheet, destHeaders);
    var masterUuidColIndex = destHeaders.indexOf('MASTER_UUID') + 1;
    var masterLinkColIndex = destHeaders.indexOf('SHEET LINK TO ROW') + 1;

    var srcUuidCol0 = indexOfHeader(header, 'MASTER_UUID');
    if (srcUuidCol0 === -1) {
        var writeCol = numCols + 1;
        sheet.getRange(1, writeCol).setValue('MASTER_UUID');
        srcUuidCol0 = writeCol - 1;
        numCols++;
        header.push('MASTER_UUID');
    }
    applyOverflowWrapStrategy_(sheet, 'MASTER_UUID');
    
    var potColIndex0 = indexOfHeader(header, 'POTENTIAL_DUPLICATES');
    if (potColIndex0 === -1) {
        var writeCol2 = numCols + 1;
        sheet.getRange(1, writeCol2).setValue('POTENTIAL_DUPLICATES');
        potColIndex0 = writeCol2 - 1;
        numCols++;
        header.push('POTENTIAL_DUPLICATES');
    }
    applyOverflowWrapStrategy_(sheet, 'POTENTIAL_DUPLICATES');


    var rowsToProcess = new Set();
    if (currentLastRow > prevLastRow) {
        for (var r = prevLastRow + 1; r <= currentLastRow; r++) rowsToProcess.add(r);
    }
    if (editedRow >= 2 && editedRow <= currentLastRow) rowsToProcess.add(editedRow);

    var sortedRows = Array.from(rowsToProcess).sort(function(a, b) { return a - b; });
    var maxHandledRow = prevLastRow;

    for (var i = 0; i < sortedRows.length; i++) {
        var rowNum = sortedRows[i];
        var rowData = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];
        if (rowData.every(function(c) { return (c === null || String(c).trim() === ''); })) {
            maxHandledRow = Math.max(maxHandledRow, rowNum);
            continue;
        }

        var first = (rowData[idxFirst] || '').toString().trim();
        var last = (rowData[idxLast] || '').toString().trim();
        var normDate = normalizeDateValue(rowData[idxDate] || new Date(), new Date().getFullYear());
        var normDob = normalizeDateValue(rowData[idxDOB], new Date().getFullYear());
        
        // --- CORE LOGIC CHANGE ---
        var srcUuid = (rowData[srcUuidCol0] || '').toString().trim();
        var masterRowFound = -1;
        var isNewRecord = false;

        if (srcUuid) {
            masterRowFound = findMasterRowByUuid(masterSheet, srcUuid, masterUuidColIndex);
            if (masterRowFound === -1) {
                // Orphan UUID, treat as new.
                srcUuid = Utilities.getUuid();
                isNewRecord = true;
            }
        } else {
            srcUuid = Utilities.getUuid();
            isNewRecord = true;
        }
        
        if (isNewRecord) {
            sheet.getRange(rowNum, srcUuidCol0 + 1).setValue(srcUuid);
        }

        var lastColLetter = columnToLetter(numCols);
        var sourceRowLink = ss.getUrl() + '#gid=' + sheetId + '&range=A' + rowNum + ':' + lastColLetter + rowNum;
        
        var outputRow = [
            normDate,
            (rowData[idxTest] || '').toString().trim(),
            first,
            last,
            normDob,
            (rowData[idxZip] || '').toString().trim(),
            spreadId, // 'SHEET' column now stores source SS ID
            'Open',
            sheet.getName(),
            ss.getName(),
            srcUuid,
            '' // Placeholder for potential duplicates
        ];

        // Perform write operation
        if (isNewRecord) {
            masterSheet.appendRow(outputRow);
            var newMasterRow = masterSheet.getLastRow();
            masterSheet.getRange(newMasterRow, masterLinkColIndex).setRichTextValue(
                SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(sourceRowLink).build()
            );
        } else {
            masterSheet.getRange(masterRowFound, 1, 1, outputRow.length).setValues([outputRow]);
            masterSheet.getRange(masterRowFound, masterLinkColIndex).setRichTextValue(
                SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(sourceRowLink).build()
            );
        }
        
        // Check for duplicates for user feedback (but don't change write logic)
        var dobIso = (normDob instanceof Date) ? Utilities.formatDate(normDob, masterTz, 'yyyy-MM-dd') : (normDob || '').toString();
        var fldKey = [first.toUpperCase(), last.toUpperCase(), dobIso].join('|');
        var matches = masterMap[fldKey] || [];
        if (matches.length > 0) {
            try { sheet.getRange(rowNum, idxDate + 1).setBackground(NEW_ROW_DATE_HIGHLIGHT); } catch (e) {}
            var potLines = matches.map(function(m) {
                return (m.fileName || '') + ' (' + (m.sheetName || '') + ') Row ' + m.row;
            });
            try { sheet.getRange(rowNum, potColIndex0 + 1).setValue(potLines.join('\n')); } catch (e) {}
        } else {
            try { sheet.getRange(rowNum, potColIndex0 + 1).clearContent(); } catch(e) {}
        }


        maxHandledRow = Math.max(maxHandledRow, rowNum);
    }

    if (maxHandledRow > prevLastRow) {
      props.setProperty(propKey, String(maxHandledRow));
    }
    
    // Final pass to mark duplicates on the master sheet
    findAndMarkDuplicates(masterSheet, destHeaders, masterTz);

  } catch (err) {
    Logger.log('onEditSourceInstallable error: ' + err.stack);
  } finally {
      if (lock) {
          lock.releaseLock();
      }
  }
}



function buildMasterNameDobMap(masterId) {
  var map = {};
  try {
    var mss = SpreadsheetApp.openById(masterId);
    var mSheet = mss.getSheetByName(DEST_SHEET_NAME);
    if (!mSheet) return map;
    var lastRow = mSheet.getLastRow();
    if (lastRow < 2) return map;

    var header = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0].map(function(h) {
      return (h || '').toString().trim().toUpperCase();
    });
    var idxFirst = indexOfHeader(header, 'FIRST NAME');
    var idxLast = indexOfHeader(header, 'LAST NAME');
    var idxDOB = indexOfHeader(header, 'DOB');
    var idxLink = indexOfHeader(header, 'SHEET LINK TO ROW');
    var idxLab = indexOfHeader(header, 'LAB (SHEET NAME)');
    var idxGroup = indexOfHeader(header, 'GROUP (FILE NAME)');
    var values = mSheet.getRange(2, 1, lastRow - 1, header.length).getValues();
    var richLinks = null;
    try {
      if (idxLink !== -1) {
        richLinks = mSheet.getRange(2, idxLink + 1, lastRow - 1, 1).getRichTextValues();
      }
    } catch (e) {
      richLinks = null;
    }

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNum = i + 2;
      var first = (row[idxFirst] || '').toString().trim().toUpperCase();
      var last = (row[idxLast] || '').toString().trim().toUpperCase();
      var dobVal = row[idxDOB];
      var dobIso = '';
      if (dobVal instanceof Date && !isNaN(dobVal.getTime())) {
        dobIso = Utilities.formatDate(dobVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (dobVal !== null && dobVal !== undefined && dobVal !== '') {
        var n = normalizeDateValue(dobVal, new Date().getFullYear());
        if (n instanceof Date && !isNaN(n.getTime())) dobIso = Utilities.formatDate(n, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        else dobIso = (n || '').toString().trim();
      } else {
        dobIso = '';
      }

      var url = '';
      if (richLinks && richLinks[i] && richLinks[i][0]) {
        try {
          var rv = richLinks[i][0];
          if (typeof rv.getLinkUrl === 'function') url = rv.getLinkUrl() || '';
        } catch (e) {
          url = '';
        }
      } else {
        url = (row[idxLink] || '').toString();
      }

      var fileName = (row[idxGroup] || '').toString();
      var sheetName = (row[idxLab] || '').toString();

      var key = [first, last, dobIso].join('|');
      if (!map[key]) map[key] = [];
      map[key].push({ fileName: fileName, sheetName: sheetName, row: rowNum, url: url });
    }
  } catch (e) {
    Logger.log('buildMasterNameDobMap error: ' + e.message);
  }
  return map;
}

function runForProvidedSheetsIncremental() {
  var result = consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME);
  Logger.log(JSON.stringify(result));
}

function consolidateSheetsIncremental(sourceSpreadsheetIds, destSpreadsheetId, destSheetName) {
  if (!Array.isArray(sourceSpreadsheetIds) || sourceSpreadsheetIds.length === 0) {
    throw new Error('Please pass an array of source spreadsheet IDs.');
  }

  var props = PropertiesService.getScriptProperties();
  var destSs = destSpreadsheetId ? SpreadsheetApp.openById(destSpreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  var tz = destSs.getSpreadsheetTimeZone();
  var destName = destSheetName || DEST_SHEET_NAME;
  var destSheet = destSs.getSheetByName(destName) || destSs.insertSheet(destName);
  var logSheet = destSs.getSheetByName(LOG_SHEET_NAME) || destSs.insertSheet(LOG_SHEET_NAME);

  var destHeaders = [
    'DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'ZIP CODE',
    'SHEET', 'SHEET LINK TO ROW', 'LAB (SHEET NAME)', 'GROUP (FILE NAME)', 'MASTER_UUID', 'POTENTIAL_DUPLICATES'
  ];
  ensureDestinationHeader(destSheet, destHeaders);
  ensureDestinationHeader(logSheet, LOG_HEADERS);
  var existingKeys = buildExistingKeys(destSheet, destHeaders, tz);
  var rowsToAppend = [];
  var linkRichTextMeta = [];
  var logsToAppend = [];
  var logRichTextMeta = [];
  var currentYear = new Date().getFullYear();

  sourceSpreadsheetIds.forEach(function(spreadId) {
    var ss, ssName, ssUrl;
    try {
      ss = SpreadsheetApp.openById(spreadId);
      ssName = ss.getName();
      ssUrl = ss.getUrl();
    } catch (e) {
      logsToAppend.push([new Date(), 'error_opening_spreadsheet: ' + e.message, spreadId, '', '', '', 'Open', '', 'system']);
      logRichTextMeta.push({ linkUrl: '', lastIndex: logsToAppend.length - 1 });
      return;
    }

    var sheets = ss.getSheets();
    sheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol === 0) return;

      var consKey = CONS_PROP_PREFIX + spreadId + '_' + sheet.getSheetId();
      var lastProcessed = Number(props.getProperty(consKey)) || 1;
      if (lastRow <= lastProcessed) {
        return;
      }

      var allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var headerRow = allValues[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });

      var sourceUuidCol0 = indexOfHeader(headerRow, 'MASTER_UUID');
      if (sourceUuidCol0 === -1) {
        sheet.getRange(1, lastCol + 1).setValue('MASTER_UUID');
        sourceUuidCol0 = lastCol;
        lastCol = sheet.getLastColumn();
        allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        headerRow = allValues[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
      }
      applyOverflowWrapStrategy_(sheet, 'MASTER_UUID');


      var desiredHeaders = ['DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'ZIP CODE'];
      var headerIndex = {};
      desiredHeaders.forEach(function(h) {
        var idx = headerRow.indexOf(h);
        if (idx === -1) {
          for (var i = 0; i < headerRow.length; i++) {
            var cand = (headerRow[i] || '').replace(/\s+/g, '');
            if (cand === h.replace(/\s+/g, '')) { idx = i; break; }
          }
        }
        headerIndex[h] = (idx === -1) ? null : (idx + 1);
      });

      for (var r = lastProcessed + 1; r <= lastRow; r++) {
        var rawRow = allValues[r - 1];
        var allTrimEmpty = rawRow.every(function(cell) {
          return (cell === null || cell === undefined || String(cell).toString().trim() === '');
        });
        if (allTrimEmpty) {
          continue;
        }

        var extracted = {};
        var anyNonEmpty = false;
        desiredHeaders.forEach(function(h) {
          var col = headerIndex[h];
          var val = '';
          if (col) {
            val = allValues[r - 1][col - 1];
            if (val !== '' && val !== null && val !== undefined) anyNonEmpty = true;
          }
          if (h === 'DATE' || h === 'DOB') {
            val = normalizeDateValue(val, currentYear);
          }
          extracted[h] = val;
        });

        var gid = sheet.getSheetId();
        var lastColLetter = columnToLetter(lastCol);
        var rowRange = 'A' + r + ':' + lastColLetter + r;
        var rowLink = ssUrl + '#gid=' + gid + '&range=' + encodeURIComponent(rowRange);

        if (!anyNonEmpty) {
          logsToAppend.push([ new Date(), 'no_data', spreadId, sheetName, r, JSON.stringify(allValues[r - 1]), 'Open', '', 'system' ]);
          logRichTextMeta.push({ linkUrl: rowLink, lastIndex: logsToAppend.length - 1 });
          continue;
        }

        var missing = [];
        MANDATORY_FIELDS.forEach(function(mf) {
          var v = extracted[mf];
          if (v === '' || v === null || v === undefined) missing.push(mf);
          if ((mf === 'DATE' || mf === 'DOB') && !(v instanceof Date)) {
            missing.push(mf + '_UNPARSED');
          }
        });

        if (missing.length > 0) {
          logsToAppend.push([ new Date(), 'missing_fields: ' + missing.join(','), spreadId, sheetName, r, JSON.stringify(extracted), 'Open', '', 'system' ]);
          logRichTextMeta.push({ linkUrl: rowLink, lastIndex: logsToAppend.length - 1 });
          continue;
        }

        var key = buildCompositeKey(extracted['DATE'], extracted['FIRST NAME'], extracted['LAST NAME'], extracted['DOB'], extracted['ZIP CODE'], tz);
        if (!key) {
          logsToAppend.push([ new Date(), 'could_not_build_key', spreadId, sheetName, r, JSON.stringify(extracted), 'Open', '', 'system' ]);
          logRichTextMeta.push({ linkUrl: rowLink, lastIndex: logsToAppend.length - 1 });
          continue;
        }
        if (existingKeys.has(key)) {
          continue;
        }

        var rowUuid = Utilities.getUuid();
        var outputRow = [
          extracted['DATE'] instanceof Date ? extracted['DATE'] : extracted['DATE'],
          extracted['TEST TYPE'] !== undefined ? extracted['TEST TYPE'] : '',
          extracted['FIRST NAME'] !== undefined ? extracted['FIRST NAME'] : '',
          extracted['LAST NAME'] !== undefined ? extracted['LAST NAME'] : '',
          extracted['DOB'] instanceof Date ? extracted['DOB'] : extracted['DOB'],
          extracted['ZIP CODE'] !== undefined ? extracted['ZIP CODE'] : '',
          spreadId,
          'Open',
          sheetName,
          ssName,
          rowUuid,
          ''
        ];
        rowsToAppend.push(outputRow);
        linkRichTextMeta.push({ linkUrl: rowLink, lastIndex: rowsToAppend.length - 1, fileName: ssName, sheetName: sheetName, noteToken: normalizeLink(rowLink) });
        try {
          sheet.getRange(r, sourceUuidCol0 + 1).setValue(rowUuid);
        } catch (wbErr) {
          Logger.log('Failed to write MASTER_UUID back to source (batch): ' + wbErr.message);
        }
        existingKeys.add(key);
      }
      try { props.setProperty(consKey, String(lastRow)); } catch (pe) { Logger.log('Could not set consolidate prop ' + consKey + ': ' + pe.message); }
    });
  });

  if (rowsToAppend.length > 0) {
    var startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    var linkColIndex = destHeaders.indexOf('SHEET LINK TO ROW') + 1;
    var uuidColIndex = destHeaders.indexOf('MASTER_UUID') + 1;
    for (var i = 0; i < linkRichTextMeta.length; i++) {
      var meta = linkRichTextMeta[i];
      var destRow = startRow + meta.lastIndex;
      var rich = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(meta.linkUrl).build();
      try {
        destSheet.getRange(destRow, linkColIndex).setRichTextValue(rich);
      } catch (e) {
        try { destSheet.getRange(destRow, linkColIndex).setValue('Open'); } catch (ee) {}
      }
      try { destSheet.getRange(destRow, linkColIndex).setNote(meta.noteToken || normalizeLink(meta.linkUrl)); } catch (ee) {}
    }
  }

  if (logsToAppend.length > 0) {
    var logStart = logSheet.getLastRow() + 1;
    logSheet.getRange(logStart, 1, logsToAppend.length, logsToAppend[0].length).setValues(logsToAppend);
    var logLinkCol = LOG_HEADERS.indexOf('LINK_TO_ROW') + 1;
    for (var j = 0; j < logRichTextMeta.length; j++) {
      var metaL = logRichTextMeta[j];
      var destRowLog = logStart + metaL.lastIndex;
      var richLog = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(metaL.linkUrl).build();
      try { logSheet.getRange(destRowLog, logLinkCol).setRichTextValue(richLog); } catch (err) {}
    }
  }

  findAndMarkDuplicates(destSheet, destHeaders, tz);
  return {
    appendedMasterRows: rowsToAppend.length,
    loggedRows: logsToAppend.length
  };
}

function findAndMarkDuplicates(destSheet, destHeaders, tz) {
  var props = PropertiesService.getScriptProperties();
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) return;

  var headerRow = destSheet.getRange(1, 1, 1, destHeaders.length).getValues()[0].map(function(h) { return (h||'').toString().trim().toUpperCase(); });
  var idxFirst = headerRow.indexOf('FIRST NAME');
  var idxLast = headerRow.indexOf('LAST NAME');
  var idxDOB = headerRow.indexOf('DOB');
  var idxSheetCol = headerRow.indexOf('SHEET');
  var idxLink = headerRow.indexOf('SHEET LINK TO ROW');
  var idxLab = headerRow.indexOf('LAB (SHEET NAME)');
  var idxGroup = headerRow.indexOf('GROUP (FILE NAME)');
  var idxPotentialDup = headerRow.indexOf('POTENTIAL_DUPLICATES');

  var dataRange = destSheet.getRange(2, 1, lastRow - 1, destHeaders.length);
  var values = dataRange.getValues();
  var richLinks = null;
  try {
    if (idxLink !== -1) richLinks = destSheet.getRange(2, idxLink + 1, lastRow - 1, 1).getRichTextValues();
  } catch (e) { richLinks = null; }

  var groups = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var sheetRowNumber = i + 2;
    var first = (row[idxFirst] || '').toString().trim().toUpperCase();
    var last = (row[idxLast] || '').toString().trim().toUpperCase();
    var dobVal = row[idxDOB];
    var dobIso = '';

    if (dobVal instanceof Date && !isNaN(dobVal.getTime())) {
      dobIso = Utilities.formatDate(dobVal, tz, 'yyyy-MM-dd');
    } else if (dobVal !== null && dobVal !== undefined && dobVal !== '') {
      var norm = normalizeDateValue(dobVal, new Date().getFullYear());
      if (norm instanceof Date && !isNaN(norm.getTime())) dobIso = Utilities.formatDate(norm, tz, 'yyyy-MM-dd');
      else dobIso = (norm || '').toString().trim();
    } else {
      dobIso = '';
    }

    var url = '';
    if (richLinks && richLinks[i] && richLinks[i][0]) {
      try {
        var richVal = richLinks[i][0];
        if (typeof richVal.getLinkUrl === 'function') url = richVal.getLinkUrl() || '';
      } catch (e) { url = ''; }
    } else {
      url = (row[idxLink] || '').toString();
    }

    var fileName = (row[idxGroup] || '').toString();
    var sheetName = (row[idxLab] || '').toString();
    var sourceId = (row[idxSheetCol] || '').toString();

    var key = [first, last, dobIso].join('|');
    if (!groups[key]) groups[key] = [];
    groups[key].push({ sheetRow: sheetRowNumber, url: url, fileName: fileName, sheetName: sheetName, sourceId: sourceId, originalValues: row });
  }

  var dupSheet = destSheet.getParent().getSheetByName(DUPLICATES_SHEET_NAME);
  if (!dupSheet) {
    dupSheet = destSheet.getParent().insertSheet(DUPLICATES_SHEET_NAME);
    var dupHeaders = ['FIRST NAME', 'LAST NAME', 'DOB', 'COUNT', 'LINKS'];
    dupSheet.getRange(1,1,1,dupHeaders.length).setValues([dupHeaders]);
  } else {
    if (dupSheet.getLastRow() >= 2) {
      dupSheet.getRange(2,1,dupSheet.getLastRow()-1, Math.max(dupSheet.getLastColumn(),5)).clearContent();
    }
    var existingDupHeader = dupSheet.getRange(1,1,1,5).getValues()[0].map(function(h){return (h||'').toString().trim().toUpperCase();});
    var expected = ['FIRST NAME','LAST NAME','DOB','COUNT','LINKS'];
    var needHeaderWrite = false;
    for (var hi = 0; hi < expected.length; hi++) {
      if (!existingDupHeader[hi] || existingDupHeader[hi] !== expected[hi]) { needHeaderWrite = true; break; }
    }
    if (needHeaderWrite) dupSheet.getRange(1,1,1,5).setValues([expected]);
  }
  applyOverflowWrapStrategy_(dupSheet, 'LINKS');

  var dupRows = [];
  var dupRichValues = [];
  var masterRichToSet = [];
  for (var k = 0; k < lastRow - 1; k++) masterRichToSet.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
  var rowsToColor = [];

  var logEntries = [];
  var logRichTextForLinks = [];

  var destSs = destSheet.getParent();
  var destUrl = destSs.getUrl();
  var destLastCol = destSheet.getLastColumn();
  var destLastColLetter = columnToLetter(destLastCol);

  for (var key in groups) {
    var arr = groups[key];
    if (arr.length <= 1) continue;
    var dupFlagKey = DUP_GROUP_PREFIX + Utilities.base64Encode(key);
    var alreadyLogged = props.getProperty(dupFlagKey);
    var shouldLogGroup = !alreadyLogged;

    var parts = key.split('|');
    var firstName = parts[0];
    var lastName = parts[1];
    var dobIsoVal = parts[2];

    var linkLabels = [];
    var linkUrls = [];
    for (var t = 0; t < arr.length; t++) {
      var label = (arr[t].fileName || '') + ' (' + (arr[t].sheetName || '') + ') Row ' + arr[t].sheetRow;
      linkLabels.push(label);
      linkUrls.push(arr[t].url || '');
    }
    var fullText = linkLabels.join(', ');
    var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
    var cursor = 0;
    for (var t = 0; t < linkLabels.length; t++) {
      var lbl = linkLabels[t];
      var start = cursor;
      var end = cursor + lbl.length;
      if (linkUrls[t]) {
        try { builder.setLinkUrl(start, end, linkUrls[t]); } catch (e) {}
      }
      cursor = end + 2;
    }

    dupRows.push([firstName, lastName, dobIsoVal, arr.length, '']);
    dupRichValues.push([builder.build()]);

    for (var a = 0; a < arr.length; a++) {
      var rowInfo = arr[a];
      var otherLabels = [];
      var otherUrls = [];
      for (var b = 0; b < arr.length; b++) {
        if (b === a) continue;
        var lbl = (arr[b].fileName || '') + ' (' + (arr[b].sheetName || '') + ') Row ' + arr[b].sheetRow;
        otherLabels.push(lbl);
        otherUrls.push(arr[b].url || '');
      }
      var dispText = otherLabels.join(', ');
      var bldr = SpreadsheetApp.newRichTextValue().setText(dispText);
      var cur = 0;
      for (var z = 0; z < otherLabels.length; z++) {
        var lblz = otherLabels[z];
        var st = cur;
        var en = cur + lblz.length;
        if (otherUrls[z]) {
          try { bldr.setLinkUrl(st, en, otherUrls[z]); } catch (e) {}
        }
        cur = en + 2;
      }
      masterRichToSet[rowInfo.sheetRow - 2] = [bldr.build()];
      rowsToColor.push(rowInfo.sheetRow);

      if (shouldLogGroup) {
        var masterRowIdx = rowInfo.sheetRow;
        var masterRowValues = rowInfo.originalValues;
        var masterLink = destUrl + '#gid=' + destSheet.getSheetId() + '&range=' + encodeURIComponent('A' + masterRowIdx + ':' + destLastColLetter + masterRowIdx);
        var dupText = otherLabels.join(', ');
        var dupBuilder = SpreadsheetApp.newRichTextValue().setText(dupText);
        var cur2 = 0;
        for (var u = 0; u < otherLabels.length; u++) {
          var lblu = otherLabels[u];
          var st2 = cur2;
          var en2 = cur2 + lblu.length;
          var urlu = otherUrls[u] || '';
          if (urlu) {
            try { dupBuilder.setLinkUrl(st2, en2, urlu); } catch (ee) {}
          }
          cur2 = en2 + 2;
        }
        logEntries.push([ new Date(), 'potential_duplicate', rowInfo.sourceId || '', rowInfo.sheetName || '', masterRowIdx, JSON.stringify(masterRowValues), 'Open', dupText, 'system' ]);
        logRichTextForLinks.push({ linkUrl: masterLink, dupRich: dupBuilder.build() });
      }
    }

    if (shouldLogGroup) {
      try { props.setProperty(dupFlagKey, String(new Date().getTime())); } catch (pe) { /* ignore */ }
    }
  }

  if (dupRows.length > 0) {
    dupSheet.getRange(2,1,dupRows.length,5).setValues(dupRows);
    try {
      dupSheet.getRange(2,5,dupRichValues.length,1).setRichTextValues(dupRichValues);
    } catch (e) {
      for (var ii = 0; ii < dupRichValues.length; ii++) {
        try { dupSheet.getRange(2 + ii, 5).setRichTextValue(dupRichValues[ii][0]); } catch (err) {}
      }
    }
  }

  if (logEntries.length > 0) {
    var logSheet = destSs.getSheetByName(LOG_SHEET_NAME) || destSs.insertSheet(LOG_SHEET_NAME);
    ensureDestinationHeader(logSheet, LOG_HEADERS);
    var logStart = logSheet.getLastRow() + 1;
    logSheet.getRange(logStart, 1, logEntries.length, logEntries[0].length).setValues(logEntries);

    var linkColIdx = LOG_HEADERS.indexOf('LINK_TO_ROW') + 1;
    var dupColIdx = LOG_HEADERS.indexOf('DUPLICATE_LINKS') + 1;
    var richVals = [];
    var dupRichArr = [];

    for (var ri = 0; ri < logRichTextForLinks.length; ri++) {
      var meta = logRichTextForLinks[ri];
      richVals.push([ SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(meta.linkUrl).build() ]);
      dupRichArr.push([ meta.dupRich || SpreadsheetApp.newRichTextValue().setText('').build() ]);
    }

    try {
      logSheet.getRange(logStart, linkColIdx, richVals.length, 1).setRichTextValues(richVals);
    } catch (e) {
      for (var q = 0; q < richVals.length; q++) {
        try { logSheet.getRange(logStart + q, linkColIdx).setRichTextValue(richVals[q][0]); } catch (err) {}
      }
    }
    try {
      logSheet.getRange(logStart, dupColIdx, dupRichArr.length, 1).setRichTextValues(dupRichArr);
    } catch (e) {
      for (var q2 = 0; q2 < dupRichArr.length; q2++) {
        try { logSheet.getRange(logStart + q2, dupColIdx).setRichTextValue(dupRichArr[q2][0]); } catch (err) {}
      }
    }
  }

  if (rowsToColor.length > 0) {
    var bgColor = NEW_ROW_DATE_HIGHLIGHT;
    for (var i2 = 0; i2 < rowsToColor.length; i2++) {
      var rr = rowsToColor[i2];
      try { destSheet.getRange(rr, 1, 1, destHeaders.length).setBackground(bgColor); } catch (e) {}
    }
    if (idxPotentialDup !== -1) {
      applyOverflowWrapStrategy_(destSheet, 'POTENTIAL_DUPLICATES');
      try {
        var richToSetRange = destSheet.getRange(2, idxPotentialDup + 1, masterRichToSet.length, 1);
        richToSetRange.setRichTextValues(masterRichToSet);
      } catch (e) {
        for (var k = 0; k < masterRichToSet.length; k++) {
          try { destSheet.getRange(2 + k, idxPotentialDup + 1).setRichTextValue(masterRichToSet[k][0]); } catch (err) {}
        }
      }
    }
  } else {
    try {
      if (idxPotentialDup !== -1) destSheet.getRange(2, idxPotentialDup + 1, Math.max(1, lastRow - 1), 1).clearContent().setBackground(null);
    } catch (e) {}
  }
}

function buildExistingKeys(destSheet, destHeaders, tz) {
  var existing = new Set();
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) return existing;
  var lastCol = destSheet.getLastColumn();
  var data = destSheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, destHeaders.length)).getValues();

  var headerRow = destSheet.getRange(1, 1, 1, destHeaders.length).getValues()[0].map(function(h){return (h||'').toString().trim().toUpperCase();});
  var idxMap = {};
  destHeaders.forEach(function(h) {
    var i = headerRow.indexOf(h);
    idxMap[h] = (i === -1) ? null : i;
  });

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    try {
      var dateVal = row[idxMap['DATE']];
      var fname = row[idxMap['FIRST NAME']];
      var lname = row[idxMap['LAST NAME']];
      var dobVal = row[idxMap['DOB']];
      var zip = row[idxMap['ZIP CODE']];

      if (!(dateVal instanceof Date)) dateVal = normalizeDateValue(dateVal, new Date().getFullYear());
      if (!(dobVal instanceof Date)) dobVal = normalizeDateValue(dobVal, new Date().getFullYear());

      var key = buildCompositeKey(dateVal, fname, lname, dobVal, zip, tz);
      if (key) existing.add(key);
    } catch (e) { continue; }
  }
  return existing;
}

function buildCompositeKey(dateVal, first, last, dobVal, zip, tz) {
  try {
    var dateStr = '';
    var dobStr = '';
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    else return null;
    if (dobVal instanceof Date && !isNaN(dobVal.getTime())) dobStr = Utilities.formatDate(dobVal, tz, 'yyyy-MM-dd');
    else return null;
    first = (first || '').toString().trim().toUpperCase();
    last = (last || '').toString().trim().toUpperCase();
    zip = (zip || '').toString().trim();
    return [dateStr, first, last, dobStr, zip].join('|');
  } catch (e) {
    return null;
  }
}

function ensureDestinationHeader(destSheet, headers) {
  var existingHeader = destSheet.getRange(1,1,1,headers.length).getValues()[0];
  var needsWrite = false;
  for (var i = 0; i < headers.length; i++) {
    if (!existingHeader[i] || existingHeader[i].toString().trim() !== headers[i]) {
      needsWrite = true;
      break;
    }
  }
  if (needsWrite) destSheet.getRange(1,1,1,headers.length).setValues([headers]);
}

function buildLinkForMatch(match) {
  try {
    if (!match) return '';
    var rawUrl = (match.url || '').toString().trim();
    var sanitizedRaw = rawUrl;
    try { sanitizedRaw = sanitizeSpreadsheetUrl ? sanitizeSpreadsheetUrl(rawUrl) : rawUrl; } catch (e) { sanitizedRaw = rawUrl; }

    var hasRangeParam = /([?&]|#)range=/.test(rawUrl) || /[A-Za-z]+\d+(?:%3A|:|%253A)[A-Za-z]*\d+/i.test(rawUrl);
    if (hasRangeParam) {
      if (sanitizedRaw) return sanitizedRaw;
      return rawUrl;
    }

    var built = '';
    try { built = buildCanonicalLinkFromMatch(match) || ''; } catch (e) { built = ''; }
    if (built) return built;

    if (sanitizedRaw) return sanitizedRaw;
    return rawUrl || '';
  } catch (err) {
    try { return (match && match.url) ? match.url : ''; } catch (e) { return ''; }
  }
}

function buildAlertsHtmlForSourceNice(alerts) {
  function escText(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function escAttr(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  var html = '<div style="font-family: Roboto, Arial, sans-serif; padding:18px; color:#1d1d1f;">';
  html += '<div style="display:flex;align-items:center;margin-bottom:12px;">';
  html += '<div style="width:8px;height:32px;background:#1e88e5;border-radius:4px;margin-right:12px;"></div>';
  html += '<h2 style="margin:0;font-size:18px;color:#1e88e5;">Potential duplicates found</h2>';
  html += '</div>';
  html += '<div style="font-size:13px;color:#3c3c3c;margin-bottom:10px;">Matches were found in <strong>Master</strong>. Click a link to open that Master row in a new tab.</div>';

  alerts.forEach(function(a) {
    html += '<div style="border:1px solid #e6f0fb;background:#f7fbff;padding:10px;border-radius:6px;margin-bottom:10px;">';
    html += '<div style="font-weight:600;margin-bottom:6px;">Source row: Row ' + escText(a.sourceRow) + ' — ' + escText(a.sourceSheetName) + '</div>';
    html += '<ul style="margin:0 0 6px 16px;padding:0;">';

    a.matches.forEach(function(m) {
      var label = escText(m.fileName || '') + ' (' + escText(m.sheetName || '') + ') Row ' + escText(m.row);
      var href = '';
      try { href = buildLinkForMatch(m) || (m.url || ''); } catch (e) { href = (m.url || ''); }

      if (href) {
        html += '<li style="margin-bottom:6px;"><a href="' + escAttr(href) + '" target="_blank" rel="noopener noreferrer" style="color:#1565c0;text-decoration:none;font-weight:500;">' + label + '</a></li>';
      } else {
        html += '<li style="margin-bottom:6px;color:#333;">' + label + '</li>';
      }
    });

    html += '</ul>';
    html += '</div>';
  });

  html += '<div style="display:flex;justify-content:flex-end;margin-top:8px;">';
  html += '<button onclick="google.script.host.close()" style="background:#1e88e5;color:white;border:none;padding:8px 12px;border-radius:6px;font-weight:600;cursor:pointer">Close</button>';
  html += '</div>';
  html += '</div>';
  return html;
}

function buildCanonicalLinkFromMatch(match) {
  try {
    if (!match) return '';
    var rawUrl = (match.url || '').toString();
    var sheetName = (match.sheetName || '').toString();
    var rowNum = Number(match.row) || null;

    var fileId = null;
    if (rawUrl) {
      var m = rawUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (m && m[1]) fileId = m[1];
    }
    if (!fileId && match.fileId) fileId = match.fileId;

    if (fileId && rowNum) {
      try {
        var targetSs = SpreadsheetApp.openById(fileId);
        var targetSheet = null;
        var sheets = targetSs.getSheets();
        for (var i = 0; i < sheets.length; i++) {
          try {
            if (sheets[i].getName() === sheetName) { targetSheet = sheets[i]; break; }
          } catch (e) {}
        }
        if (!targetSheet && sheetName) {
          var lname = sheetName.toLowerCase();
          for (var j = 0; j < sheets.length; j++) {
            try {
              if (sheets[j].getName().toLowerCase() === lname) { targetSheet = sheets[j]; break; }
            } catch (e) {}
          }
        }
        if (targetSheet) {
          var gid = targetSheet.getSheetId();
          var lastCol = Math.max(1, targetSheet.getLastColumn() || 1);
          var lastColLetter = columnToLetter(lastCol);
          var a1 = 'A' + rowNum + ':' + lastColLetter + rowNum;
          return 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit#gid=' + gid + '&range=' + encodeURIComponent(a1);
        }
      } catch (openErr) {
      }
    }

    if (rawUrl) {
      var s = rawUrl.toString().trim();
      s = s.replace(/[\u200B-\u200F\u2028-\u202F\uFEFF\u2060-\u206F\u00A0]/g, '');
      s = s.replace(/\uFFFD/g, '');
      for (var k = 0; k < 4; k++) {
        try {
          var dec = decodeURIComponent(s);
          if (dec === s) break;
          s = dec;
        } catch (e) { break; }
      }
      var idM = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
      var gidM = s.match(/[?&]gid=(\d+)/) || s.match(/#gid=(\d+)/);
      var fId = idM ? idM[1] : null;
      var gId = gidM ? gidM[1] : null;
      if (fId && gId) return 'https://docs.google.com/spreadsheets/d/' + fId + '/edit#gid=' + gId;
      if (fId) return 'https://docs.google.com/spreadsheets/d/' + fId + '/edit';
      return s;
    }
    return '';
  } catch (err) {
    return (match && match.url) ? match.url : '';
  }
}

/** -------- Menu handlers (thin wrappers with confirmation & toasts) -------- */
function ui_rebuildMasterFullReset_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Rebuild Master (Full Reset)',
    'This will CLEAR rows in Master, Log, and Duplicates (headers/formatting kept), ' +
    'reset cursors & duplicate logs, then reimport everything from sources. Proceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    SpreadsheetApp.getActive().toast('Starting full reset & reimport…', 'Data Consolidator', 5);
    resetMasterAndReimportAll();
    SpreadsheetApp.getActive().toast('Full rebuild complete ✅', 'Data Consolidator', 5);
    SpreadsheetApp.flush();
  } catch (e) {
    ui.alert('Full Reset Failed', e.message, ui.ButtonSet.OK);
  }
}

function ui_reimportHistoryNoClear_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Reimport History (Keep Master)',
    'This will reset incremental cursors (and duplicate-group memory) and re-run import without clearing Master. Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    SpreadsheetApp.getActive().toast('Reimporting history (keeping Master)…', 'Data Consolidator', 5);
    resetCursorsAndReimportWithoutClearing();
    SpreadsheetApp.getActive().toast('Reimport complete ✅', 'Data Consolidator', 5);
    SpreadsheetApp.flush();
  } catch (e) {
    ui.alert('Reimport Failed', e.message, ui.ButtonSet.OK);
  }
}

function ui_initSourceLastRows_() {
  try {
    SpreadsheetApp.getActive().toast('Seeding source last rows…', 'Data Consolidator', 3);
    initSourceLastRows();
    SpreadsheetApp.getActive().toast('Seeded source last rows ✅', 'Data Consolidator', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('initSourceLastRows failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_initConsolidateLastRows_() {
  try {
    SpreadsheetApp.getActive().toast('Seeding consolidation cursors…', 'Data Consolidator', 3);
    initConsolidateLastRows();
    SpreadsheetApp.getActive().toast('Seeded consolidation cursors ✅', 'Data Consolidator', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('initConsolidateLastRows failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_createTriggersForSources_() {
  try {
    SpreadsheetApp.getActive().toast('Creating onEdit triggers for sources…', 'Data Consolidator', 5);
    createTriggersForSources();
    SpreadsheetApp.getActive().toast('Source triggers created ✅', 'Data Consolidator', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('createTriggersForSources failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_createDailyConsolidationTrigger_() {
  try {
    SpreadsheetApp.getActive().toast('Creating daily 2am consolidation trigger…', 'Data Consolidator', 5);
    createDailyConsolidationTrigger();
    SpreadsheetApp.getActive().toast('Daily trigger created ✅', 'Data Consolidator', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('createDailyConsolidationTrigger failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_openLogSheet_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
    ss.setActiveSheet(sh);
    SpreadsheetApp.getActive().toast('Opened Log sheet', 'Data Consolidator', 2);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Open Log failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets a sheet by name or creates it with specified headers if it doesn't exist.
 * This is a helper needed by the sidebar functions.
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