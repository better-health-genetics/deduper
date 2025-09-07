/** ===========================================================================
 * 
 *                       HELPER FUNCTIONS
 * 
 * ===========================================================================/

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

/**
 * Normalizes a Google Sheet URL to its canonical form (base URL + gid).
 * @param {string} url The URL to normalize.
 * @return {string} The normalized URL.
 */
function normalizeLink(url) {
  if (!url || typeof url !== 'string') return '';
  try {
    var decodedUrl = url;
    // Attempt to decode URI components multiple times in case of double encoding
    for (var i = 0; i < 3; i++) {
      var nextDecoded = decodeURIComponent(decodedUrl);
      if (nextDecoded === decodedUrl) break;
      decodedUrl = nextDecoded;
    }
    
    var idMatch = decodedUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    var gidMatch = decodedUrl.match(/[?&#]gid=(\d+)/);
    
    var fileId = idMatch ? idMatch[1] : null;
    var gid = gidMatch ? gidMatch[1] : null;
    
    if (fileId && gid) {
      return 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit#gid=' + gid;
    } else if (fileId) {
      return 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit';
    }
    return url;
  } catch (e) {
    return url; // Return original on error
  }
}

// Define sanitizeSpreadsheetUrl to avoid another potential error, using the same logic.
var sanitizeSpreadsheetUrl = normalizeLink;

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
