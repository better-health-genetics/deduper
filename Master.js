
/****************************************************************************************************
 * ___  ____ ___ ___ ____ ____    _  _ ____ ____ _    ___ _  _    ____ ____ _  _ ____ ___ _ ____ ____ 
 * |__] |___  |   |  |___ |__/    |__| |___ |__| |     |  |__|    | __ |___ |\ | |___  |  | |    [__  
 * |__] |___  |   |  |___ |  \    |  | |___ |  | |___  |  |  |    |__] |___ | \| |___  |  | |___ ___] 
 * 
 *                   BETTER HEALTH GENETICS — ADMIN & LIBRARY SCRIPT
 ****************************************************************************************************
 * ARCHITECTURE OVERVIEW
 * ---------------------
 * This script serves two primary roles:
 * 1. As the main backend for the MASTER spreadsheet, providing an "Admin Overview" sidebar and the core data consolidation logic.
 * 2. As a LIBRARY, providing a shared UI and data processing logic for all SOURCE spreadsheets.
 ****************************************************************************************************/

/* ---------- CONFIG (SHARED) ---------- */
var SOURCE_IDS = [
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
var MASTER_SPREADSHEET_ID = '161uw5s1lOwhV7YTX8uLKDw6TMl_QLEpSqxmHEftZzVA';
var DEST_SHEET_NAME = 'Master';
var LOG_SHEET_NAME = 'Log';
var DUPLICATES_SHEET_NAME = 'Duplicates';
var HELPER_SHEET_NAME = '_QueryHelper';
var MANDATORY_FIELDS = ['DATE', 'FIRST NAME', 'LAST NAME', 'DOB'];
var NEW_ROW_DATE_HIGHLIGHT = '#ffcccc';
var PROP_PREFIX = 'lastRow_';
var CONS_PROP_PREFIX = 'consolidate_lastRow_';
var DUP_GROUP_PREFIX = 'dup_logged_';
var LOG_HEADERS = ['TIMESTAMP', 'REASON', 'SOURCE_SPREADSHEET_ID', 'SOURCE_SHEET_NAME', 'ROW_NUMBER', 'ORIGINAL_VALUES', 'LINK_TO_ROW', 'DUPLICATE_LINKS', 'TRIGGERING_USER'];


/** ===========================================================================
 * 
 *                         CORE HELPER FUNCTIONS
 *      (Moved to top to prevent ReferenceError issues in V8 runtime)
 * 
 * ===========================================================================/

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

function ensureDestinationHeader(destSheet, headers) {
  if (!destSheet) return;
  var lastCol = Math.max(headers.length, destSheet.getLastColumn());
  if (lastCol < 1) lastCol = headers.length;
  var existingHeader = destSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var needsWrite = false;
  for (var i = 0; i < headers.length; i++) {
    if (!existingHeader[i] || existingHeader[i].toString().trim() !== headers[i]) {
      needsWrite = true;
      break;
    }
  }
  if (needsWrite) {
    destSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    destSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

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
      }
    }
  } catch (e) {
    Logger.log('Failed to apply wrap strategy for header "' + headerName + '" on sheet "' + sheet.getName() + '": ' + e.message);
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
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (!value) return '';
  var s = value.toString().trim();
  var today = new Date();
  defaultYear = defaultYear || today.getFullYear();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    var d = new Date(defaultYear, parseInt(m[1], 10) - 1, parseInt(m[2], 10));
    if (d.getTime() > today.getTime()) d.setFullYear(d.getFullYear() - 1);
    return d;
  }
  var parsed = new Date(s);
  return !isNaN(parsed.getTime()) ? parsed : s;
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


/** ===========================================================================
 * 
 *                  MASTER SHEET UI & ADMIN FUNCTIONS
 * 
 * ===========================================================================/

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Data Consolidator')
      .addItem('Show Admin Overview', 'showAdminOverviewSidebar')
      .addSeparator()
      .addItem('🔄 Rebuild Master (Full Reset)', 'ui_rebuildMasterFullReset_')
      .addItem('📥 Reimport History (Keep Master)', 'ui_reimportHistoryNoClear_')
      .addSeparator()
      .addItem('⚙️ Create Source onEdit Triggers', 'ui_createTriggersForSources_')
      .addItem('⏰ Create Daily Consolidation (2am)', 'ui_createDailyConsolidationTrigger_')
      .addSeparator()
      .addItem('📜 Open Log Sheet', 'ui_openLogSheet_')
      .addToUi();
  } catch (e) {
    Logger.log('addConsolidatorMenu error: ' + e.message);
  }
  showAdminOverviewSidebar();
}

function showAdminOverviewSidebar() {
  var tpl = HtmlService.createTemplateFromFile('index');
  tpl.appMode = 'admin';
  var html = tpl.evaluate()
    .setTitle('Admin Overview')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAdminOverviewData() {
  var allStats = [];
  var masterMap = buildMasterNameDobMap(MASTER_SPREADSHEET_ID);

  SOURCE_IDS.forEach(function(id, index) {
    var sourceSs;
    var sheetName = 'Unknown Sheet';
    try {
      sourceSs = SpreadsheetApp.openById(id);
      sheetName = sourceSs.getName();

      var stats = { id: id, name: sheetName, totalChecked: 0, totalDuplicates: 0, percentage: 0 };
      var sheets = sourceSs.getSheets();

      sheets.forEach(function(sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
        var idxFirst = indexOfHeader(headers, 'FIRST NAME');
        var idxLast = indexOfHeader(headers, 'LAST NAME');
        var idxDOB = indexOfHeader(headers, 'DOB');
        var idxDate = indexOfHeader(headers, 'DATE');
        if (idxFirst === -1 || idxLast === -1 || idxDOB === -1 || idxDate === -1) return;
        
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const dayOfWeek = today.getDay();
        const daysSinceMonday = (dayOfWeek + 6) % 7;
        const lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues().forEach(function(row) {
            var entryDate = normalizeDateValue(row[idxDate], new Date().getFullYear());
            if (entryDate instanceof Date && entryDate >= lastMonday) {
                stats.totalChecked++;
                var first = (row[idxFirst] || '').toString().trim().toUpperCase();
                var last = (row[idxLast] || '').toString().trim().toUpperCase();
                var dob = normalizeDateValue(row[idxDOB], new Date().getFullYear());
                var dobIso = (dob instanceof Date) ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
                var key = [first, last, dobIso].join('|');
                if (masterMap[key] && masterMap[key].length > 0) {
                    stats.totalDuplicates++;
                }
            }
        });
      });

      stats.percentage = stats.totalChecked > 0 ? Math.round((stats.totalDuplicates / stats.totalChecked) * 100) : 0;
      allStats.push(stats);

    } catch (e) {
      Logger.log('Could not process sheet ' + id + ': ' + e.message);
      allStats.push({ id: id, name: `Sheet ${index+1} (Error)`, error: e.message, percentage: 0, totalChecked: 0, totalDuplicates: 0 });
    }
  });
  return allStats;
}

/** ===========================================================================
 * 
 *                     FUNCTIONS EXPOSED TO THE LIBRARY
 * 
 * ===========================================================================/

function getSidebarUiForLibrary() {
  var tpl = HtmlService.createTemplateFromFile('index');
  tpl.appMode = 'entry';
  return tpl.evaluate().setTitle('BHG DeDuper').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getHealthDataForLibrary() {
    try {
        const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
        const masterMap = buildMasterNameDobMap(MASTER_SPREADSHEET_ID);
        var stats = { totalChecked: 0, totalDuplicates: 0 };

        const today = new Date(); today.setHours(0, 0, 0, 0);
        const dayOfWeek = today.getDay();
        const daysSinceMonday = (dayOfWeek + 6) % 7;
        const lastMonday = new Date(today.getTime() - daysSinceMonday * 24 * 60 * 60 * 1000);

        sourceSs.getSheets().forEach(function(sheet) {
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return;
            var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
            var idxFirst = indexOfHeader(headers, 'FIRST NAME');
            var idxLast = indexOfHeader(headers, 'LAST NAME');
            var idxDOB = indexOfHeader(headers, 'DOB');
            var idxDate = indexOfHeader(headers, 'DATE');
            if (idxFirst === -1 || idxLast === -1 || idxDOB === -1 || idxDate === -1) return;

            sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues().forEach(function(row) {
                var entryDate = normalizeDateValue(row[idxDate], new Date().getFullYear());
                if (entryDate instanceof Date && entryDate >= lastMonday) {
                    stats.totalChecked++;
                    var first = (row[idxFirst] || '').toString().trim().toUpperCase();
                    var last = (row[idxLast] || '').toString().trim().toUpperCase();
                    var dob = normalizeDateValue(row[idxDOB], new Date().getFullYear());
                    var dobIso = (dob instanceof Date) ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
                    var key = [first, last, dobIso].join('|');
                    if (masterMap[key] && masterMap[key].length > 0) {
                        stats.totalDuplicates++;
                    }
                }
            });
        });

        const percentage = stats.totalChecked > 0 ? Math.round((stats.totalDuplicates / stats.totalChecked) * 100) : 0;
        return { percentage: percentage, totalChecked: stats.totalChecked, totalDuplicates: stats.totalDuplicates };
    } catch (e) {
        Logger.log('Error in getHealthDataForLibrary: ' + e.stack);
        return { error: 'Failed to calculate health data: ' + e.message, percentage: 0, totalChecked: 0, totalDuplicates: 0 };
    }
}

function processSourceSheetRecordForLibrary(formData) {
  try {
    const { firstName, lastName, dob: dobString } = formData;
    if (!firstName || !lastName || !dobString) {
      return { success: false, message: 'Error: All fields are required.' };
    }
    const duplicatesFound = queryForDuplicates_(firstName, lastName, dobString);
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const [year, month, day] = dobString.split('-').map(Number);
    const dobForSheet = new Date(year, month - 1, day);
    
    // Assumes simple [DATE, '', FNAME, LNAME, DOB] schema. Adjust columns as needed.
    const newRowData = [new Date(), '', firstName, lastName, dobForSheet];
    activeSheet.appendRow(newRowData);
    
    if (duplicatesFound > 0) {
      activeSheet.getRange(activeSheet.getLastRow(), 1, 1, newRowData.length).setBackground(NEW_ROW_DATE_HIGHLIGHT);
      return { success: false, message: `${duplicatesFound} potential duplicate(s) found in Master.`, duplicatesFound: duplicatesFound };
    } else {
      return { success: true, message: 'Record added. No duplicates found in Master.' };
    }
  } catch (e) {
    Logger.log('Error in processSourceSheetRecordForLibrary: ' + e.stack);
    return { success: false, message: 'An unexpected server error occurred.' };
  }
}

function queryForDuplicates_(firstName, lastName, dobString) {
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
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


/** ===========================================================================
 * 
 *                  ORIGINAL DATA CONSOLIDATOR SCRIPT
 * 
 * ===========================================================================/

function clearSheetKeepHeader_(sheet, minCols) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(minCols || sheet.getLastColumn(), 1);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent().clearNote().setBackground(null);
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

  var master = getOrCreateSheet(destSs, DEST_SHEET_NAME, expectedHeaders);
  var log    = getOrCreateSheet(destSs, LOG_SHEET_NAME, LOG_HEADERS);
  var dups   = getOrCreateSheet(destSs, DUPLICATES_SHEET_NAME, ['FIRST NAME','LAST NAME','DOB','COUNT','LINKS']);

  try {
    var f = master.getFilter && master.getFilter();
    if (f) f.remove();
    var curCols = master.getLastColumn();
    if (curCols > expectedHeaders.length) {
      master.deleteColumns(expectedHeaders.length + 1, curCols - expectedHeaders.length);
    }
    clearSheetKeepHeader_(master, expectedHeaders.length);
  } catch (e) {
    Logger.log('Master cleanup failed: ' + e.message);
  }

  applyOverflowWrapStrategy_(master, 'POTENTIAL_DUPLICATES');
  applyOverflowWrapStrategy_(master, 'MASTER_UUID');
  
  clearSheetKeepHeader_(log);
  
  clearSheetKeepHeader_(dups);
  applyOverflowWrapStrategy_(dups, 'LINKS');

  resetConsolidationCursors_();
  resetDuplicateGroupLog_();
  
  consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME);

  try {
    var tz = destSs.getSpreadsheetTimeZone();
    findAndMarkDuplicates(master, expectedHeaders, tz);
  } catch (e) {
    Logger.log('Post-rebuild duplicate pass failed: ' + e.message);
  }
}

function shouldDebounceRow_(spreadId, sheetId, rowNum, windowMs) {
  try {
    var key = 'debounce_' + spreadId + '_' + sheetId + '_' + rowNum;
    var cache = CacheService.getScriptCache();
    if (cache.get(key)) return true;
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
}

function normalizeLink(url) {
  if (!url) return '';
  var s = String(url || '').trim();
  try { s = decodeURIComponent(s); } catch (e) {}
  var idM = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  var fileId = idM ? idM[1] : '';
  var gidM = s.match(/[?&]gid=(\d+)|#gid=(\d+)/);
  var gid = gidM ? (gidM[1] || gidM[2]) : '';
  var rM = s.match(/range=[A-Za-z]+(\d+)/i);
  var row = rM ? rM[1] : '';
  return [fileId, gid, row].join('|');
}

function findMasterRowBySourceLink(masterSheet, sourceRowLink, linkColIndex) {
  if (!masterSheet || !sourceRowLink) return -1;
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return -1;
  var normalizedTarget = normalizeLink(sourceRowLink);
  var richVals = masterSheet.getRange(2, linkColIndex, lastRow - 1, 1).getRichTextValues();
  for (var i = 0; i < richVals.length; i++) {
    var url = richVals[i][0].getLinkUrl();
    if (normalizeLink(url) === normalizedTarget) return 2 + i;
  }
  return -1;
}

function findMasterRowByUuid(masterSheet, uuid, uuidColIndex) {
  if (!masterSheet || !uuid) return -1;
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = masterSheet.getRange(2, uuidColIndex, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if ((vals[i][0] || '').toString().trim() === uuid) return 2 + i;
  }
  return -1;
}

function initSourceLastRows() {
  var props = PropertiesService.getScriptProperties();
  SOURCE_IDS.forEach(function(id) {
    try {
      var ss = SpreadsheetApp.openById(id);
      ss.getSheets().forEach(function(sh) {
        var key = PROP_PREFIX + id + '_' + sh.getSheetId();
        props.setProperty(key, String(sh.getLastRow()));
      });
    } catch(e) { Logger.log("Error in initSourceLastRows for " + id + ": " + e.message); }
  });
}

function initConsolidateLastRows() {
  var props = PropertiesService.getScriptProperties();
  SOURCE_IDS.forEach(function(id) {
    try {
      var ss = SpreadsheetApp.openById(id);
      ss.getSheets().forEach(function(sh) {
        var key = CONS_PROP_PREFIX + id + '_' + sh.getSheetId();
        props.setProperty(key, '1');
      });
    } catch(e) { Logger.log("Error in initConsolidateLastRows for " + id + ": " + e.message); }
  });
}

function createTriggersForSources() {
  SOURCE_IDS.forEach(function(spreadId) {
    try {
      ScriptApp.getProjectTriggers().forEach(function(t) {
        if (t.getHandlerFunction() === 'onEditSourceInstallable' && t.getTriggerSourceId() === spreadId) {
          ScriptApp.deleteTrigger(t);
        }
      });
      ScriptApp.newTrigger('onEditSourceInstallable').forSpreadsheet(spreadId).onEdit().create();
      Logger.log('Created trigger for: ' + spreadId);
    } catch (err) {
      Logger.log('Failed to create trigger for ' + spreadId + ': ' + err.message);
    }
  });
}

function createDailyConsolidationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runForProvidedSheetsIncremental') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runForProvidedSheetsIncremental').timeBased().everyDays(1).atHour(2).create();
}

function onEditSourceInstallable(e) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;
  try {
    var ss = e.source;
    var sheet = e.range.getSheet();
    var spreadId = ss.getId();
    var editedRow = e.range.getRow();
    if (editedRow === 1 || shouldDebounceRow_(spreadId, sheet.getSheetId(), editedRow, 3000)) return;

    var headerUpper = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){return (h||'').toString().trim().toUpperCase();});
    if (!isRelevantColumn_(headerUpper, e.range.getColumn())) return;

    var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    var masterTz = masterSs.getSpreadsheetTimeZone();
    var destHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(function(h){return (h||'').toString().trim().toUpperCase()});

    var potColIndex0 = indexOfHeader(headerUpper, 'POTENTIAL_DUPLICATES');
    if (potColIndex0 === -1) { potColIndex0 = sheet.getLastColumn(); sheet.getRange(1, potColIndex0 + 1).setValue('POTENTIAL_DUPLICATES'); }
    applyOverflowWrapStrategy_(sheet, 'POTENTIAL_DUPLICATES');

    var srcUuidCol0 = indexOfHeader(headerUpper, 'MASTER_UUID');
    if (srcUuidCol0 === -1) { srcUuidCol0 = sheet.getLastColumn(); sheet.getRange(1, srcUuidCol0 + 1).setValue('MASTER_UUID'); }
    applyOverflowWrapStrategy_(sheet, 'MASTER_UUID');

    var row = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    var srcUuid = (row[srcUuidCol0] || '').toString().trim();
    if (!srcUuid) {
      srcUuid = Utilities.getUuid();
      sheet.getRange(editedRow, srcUuidCol0 + 1).setValue(srcUuid);
    }

    var outputRow = [
      normalizeDateValue(row[indexOfHeader(headerUpper, 'DATE')]), row[indexOfHeader(headerUpper, 'TEST TYPE')] || '',
      row[indexOfHeader(headerUpper, 'FIRST NAME')] || '', row[indexOfHeader(headerUpper, 'LAST NAME')] || '',
      normalizeDateValue(row[indexOfHeader(headerUpper, 'DOB')]), row[indexOfHeader(headerUpper, 'ZIP CODE')] || '',
      spreadId, 'Open', sheet.getName(), ss.getName(), srcUuid, ''
    ];
    
    var masterUuidColIndex = destHeaders.indexOf('MASTER_UUID') + 1;
    var masterRowByUuid = findMasterRowByUuid(masterSheet, srcUuid, masterUuidColIndex);

    if (masterRowByUuid !== -1) {
      masterSheet.getRange(masterRowByUuid, 1, 1, outputRow.length).setValues([outputRow]);
    } else {
      masterSheet.appendRow(outputRow);
      masterRowByUuid = masterSheet.getLastRow();
    }
    
    var sourceRowLink = ss.getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + editedRow;
    var masterLinkColIndex = destHeaders.indexOf('SHEET LINK TO ROW') + 1;
    masterSheet.getRange(masterRowByUuid, masterLinkColIndex).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(sourceRowLink).build());
    
    findAndMarkDuplicates(masterSheet, destHeaders.map(String), masterTz);
  } finally {
    lock.releaseLock();
  }
}

function buildMasterNameDobMap(masterId) {
  var map = {};
  var mss = SpreadsheetApp.openById(masterId);
  var mSheet = mss.getSheetByName(DEST_SHEET_NAME);
  if (!mSheet || mSheet.getLastRow() < 2) return map;
  
  var header = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });
  var idxFirst = indexOfHeader(header, 'FIRST NAME');
  var idxLast = indexOfHeader(header, 'LAST NAME');
  var idxDOB = indexOfHeader(header, 'DOB');
  
  var values = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, mSheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    var dob = normalizeDateValue(values[i][idxDOB]);
    var dobIso = (dob instanceof Date) ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    var key = [(values[i][idxFirst] || '').toUpperCase(), (values[i][idxLast] || '').toUpperCase(), dobIso].join('|');
    if (!map[key]) map[key] = [];
    map[key].push({row: i + 2});
  }
  return map;
}

function runForProvidedSheetsIncremental() { consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME); }

function consolidateSheetsIncremental(sourceSpreadsheetIds, destSpreadsheetId, destSheetName) {
  var props = PropertiesService.getScriptProperties();
  var destSs = SpreadsheetApp.openById(destSpreadsheetId);
  var tz = destSs.getSpreadsheetTimeZone();
  var destSheet = getOrCreateSheet(destSs, destSheetName, ['DATE', 'TEST TYPE', 'FIRST NAME', 'LAST NAME', 'DOB', 'ZIP CODE', 'SHEET', 'SHEET LINK TO ROW', 'LAB (SHEET NAME)', 'GROUP (FILE NAME)', 'MASTER_UUID', 'POTENTIAL_DUPLICATES']);
  
  sourceSpreadsheetIds.forEach(function(id) {
    var ss = SpreadsheetApp.openById(id);
    ss.getSheets().forEach(function(sheet) {
      var consKey = CONS_PROP_PREFIX + id + '_' + sheet.getSheetId();
      var lastProcessed = Number(props.getProperty(consKey)) || 1;
      var lastRow = sheet.getLastRow();
      if (lastRow <= lastProcessed) return;

      var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){return (h||'').toString().trim().toUpperCase()});
      var srcUuidCol0 = indexOfHeader(header, 'MASTER_UUID');
      if (srcUuidCol0 === -1) { srcUuidCol0 = sheet.getLastColumn(); sheet.getRange(1, srcUuidCol0 + 1).setValue('MASTER_UUID'); applyOverflowWrapStrategy_(sheet, 'MASTER_UUID'); }

      var data = sheet.getRange(lastProcessed + 1, 1, lastRow - lastProcessed, sheet.getLastColumn()).getValues();
      var rowsToAppend = [];
      for(var i=0; i < data.length; i++) {
        var row = data[i];
        var uuid = Utilities.getUuid();
        sheet.getRange(lastProcessed + 1 + i, srcUuidCol0 + 1).setValue(uuid);
        rowsToAppend.push([
          normalizeDateValue(row[indexOfHeader(header, 'DATE')]), row[indexOfHeader(header, 'TEST TYPE')] || '',
          row[indexOfHeader(header, 'FIRST NAME')] || '', row[indexOfHeader(header, 'LAST NAME')] || '',
          normalizeDateValue(row[indexOfHeader(header, 'DOB')]), row[indexOfHeader(header, 'ZIP CODE')] || '',
          id, 'Open', sheet.getName(), ss.getName(), uuid, ''
        ]);
      }
      if(rowsToAppend.length > 0) {
        var startRow = destSheet.getLastRow() + 1;
        destSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      }
      props.setProperty(consKey, String(lastRow));
    });
  });
  findAndMarkDuplicates(destSheet, destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0], tz);
}

function findAndMarkDuplicates(destSheet, destHeaders, tz) {
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) return;

  var data = destSheet.getRange(2, 1, lastRow - 1, destSheet.getLastColumn()).getValues();
  var groups = {};
  var headerUpper = destHeaders.map(function(h){return (h||'').toString().trim().toUpperCase()});
  var idxFirst = indexOfHeader(headerUpper, 'FIRST NAME');
  var idxLast = indexOfHeader(headerUpper, 'LAST NAME');
  var idxDOB = indexOfHeader(headerUpper, 'DOB');
  var potDupIdx = indexOfHeader(headerUpper, 'POTENTIAL_DUPLICATES');
  
  for (var i = 0; i < data.length; i++) {
    var dob = normalizeDateValue(data[i][idxDOB]);
    var dobIso = (dob instanceof Date) ? Utilities.formatDate(dob, tz, 'yyyy-MM-dd') : '';
    var key = [(data[i][idxFirst]||'').toUpperCase(), (data[i][idxLast]||'').toUpperCase(), dobIso].join('|');
    if (!groups[key]) groups[key] = [];
    groups[key].push({row: i + 2});
  }
  
  var potDupVals = [];
  for(var i=0; i<data.length; i++) potDupVals.push(['']);

  for (var key in groups) {
    if (groups[key].length > 1) {
      var rows = groups[key].map(function(item) { return item.row; });
      var links = rows.map(function(r) { return 'Master Row ' + r; }).join(', ');
      rows.forEach(function(rowNum) {
        potDupVals[rowNum - 2] = [links];
        destSheet.getRange(rowNum, 1, 1, destSheet.getLastColumn()).setBackground(NEW_ROW_DATE_HIGHLIGHT);
      });
    }
  }

  if (potDupIdx !== -1 && potDupVals.length > 0) {
    destSheet.getRange(2, potDupIdx + 1, potDupVals.length, 1).setValues(potDupVals);
    applyOverflowWrapStrategy_(destSheet, 'POTENTIAL_DUPLICATES');
  }
}

function ui_rebuildMasterFullReset_() { SpreadsheetApp.getUi().alert('Full reset started...'); resetMasterAndReimportAll(); SpreadsheetApp.getUi().alert('Full reset complete.'); }
function ui_reimportHistoryNoClear_() { SpreadsheetApp.getUi().alert('Reimport started...'); resetCursorsAndReimportWithoutClearing(); SpreadsheetApp.getUi().alert('Reimport complete.'); }
function ui_initSourceLastRows_() { initSourceLastRows(); SpreadsheetApp.getUi().alert('Source rows seeded.'); }
function ui_initConsolidateLastRows_() { initConsolidateLastRows(); SpreadsheetApp.getUi().alert('Consolidation cursors seeded.'); }
function ui_createTriggersForSources_() { createTriggersForSources(); SpreadsheetApp.getUi().alert('Triggers created.'); }
function ui_createDailyConsolidationTrigger_() { createDailyConsolidationTrigger(); SpreadsheetApp.getUi().alert('Daily trigger created.'); }
function ui_openLogSheet_() { var ss = SpreadsheetApp.getActive(); ss.setActiveSheet(ss.getSheetByName(LOG_SHEET_NAME)); }

function buildCanonicalLinkFromMatch(match) {
  // Full implementation from user's provided script
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