/** ===========================================================================
 * 
 *                  DATA CONSOLIDATOR SCRIPT
 * 
 * ===========================================================================/

/**
 * ** OVERHAULED onEdit TRIGGER **
 * This function now determines the edited rows directly from the event object (`e.range`),
 * making it more reliable for single edits, pastes, and deletions. It prioritizes MASTER_UUID
 * for updates. If a UUID is present, it updates the Master record. If not, it creates a new record.
 */
function onEditSourceInstallable(e) {
  var lock;
  try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(15000)) { // Increased lock wait time
        Logger.log('onEditSourceInstallable could not get lock for ' + e.source.getId());
        return;
    }
    
    var ss = e.source;
    var sheet = e.range.getSheet();
    var spreadId = ss.getId();
    var sheetId = sheet.getSheetId();
    var editedCol = e.range.getColumn();

    var numCols = Math.max(1, sheet.getLastColumn());
    var headerUpper = sheet.getRange(1, 1, 1, numCols).getValues()[0].map(function(h) { return (h || '').toString().trim().toUpperCase(); });

    if (!isRelevantColumn_(headerUpper, editedCol)) return;

    var masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    var masterSheet = masterSs.getSheetByName(DEST_SHEET_NAME);
    var masterTz = masterSs.getSpreadsheetTimeZone();
    var masterMap = buildMasterNameDobMap(MASTER_SPREADSHEET_ID); // For duplicate suggestions

    var destHeaders = ['DATE','TEST TYPE','FIRST NAME','LAST NAME','DOB','ZIP CODE','SHEET','SHEET LINK TO ROW','LAB (SHEET NAME)','GROUP (FILE NAME)','MASTER_UUID','POTENTIAL_DUPLICATES'];
    ensureDestinationHeader(masterSheet, destHeaders);
    var masterUuidColIndex = destHeaders.indexOf('MASTER_UUID') + 1;
    var masterLinkColIndex = destHeaders.indexOf('SHEET LINK TO ROW') + 1;
    var masterDateColIndex = destHeaders.indexOf('DATE') + 1;
    var masterDobColIndex = destHeaders.indexOf('DOB') + 1;


    var idxFirst = indexOfHeader(headerUpper, 'FIRST NAME');
    var idxLast = indexOfHeader(headerUpper, 'LAST NAME');
    var idxDOB = indexOfHeader(headerUpper, 'DOB');
    var idxDate = indexOfHeader(headerUpper, 'DATE');
    var idxZip = indexOfHeader(headerUpper, 'ZIP CODE');
    var idxTest = indexOfHeader(headerUpper, 'TEST TYPE');
    if (idxFirst === -1 || idxLast === -1 || idxDOB === -1) return;


    var srcUuidCol0 = indexOfHeader(headerUpper, 'MASTER_UUID');
    if (srcUuidCol0 === -1) {
        srcUuidCol0 = ensureHeaderAndReturnIndex(sheet, 'MASTER_UUID');
        applyOverflowWrapStrategy_(sheet, 'MASTER_UUID');
        numCols = sheet.getLastColumn();
    }
    var potColIndex0 = indexOfHeader(headerUpper, 'POTENTIAL_DUPLICATES');
    if (potColIndex0 === -1) {
        potColIndex0 = ensureHeaderAndReturnIndex(sheet, 'POTENTIAL_DUPLICATES');
        applyOverflowWrapStrategy_(sheet, 'POTENTIAL_DUPLICATES');
        numCols = sheet.getLastColumn();
    }

    // --- REFACTORED ROW PROCESSING ---
    var startRow = e.range.getRow();
    var numRowsInRange = e.range.getNumRows();
    var rowsToProcess = [];
    for (var i = 0; i < numRowsInRange; i++) {
        var rowNum = startRow + i;
        if (rowNum > 1) rowsToProcess.push(rowNum);
    }
    var sortedRows = Array.from(new Set(rowsToProcess)).sort(function(a, b) { return a - b; });
    if (sortedRows.length === 0) return;
    // --- END REFACTOR ---
    
    var dataToRead = sheet.getRange(sortedRows[0], 1, sortedRows[sortedRows.length - 1] - sortedRows[0] + 1, numCols).getValues();

    for (var i = 0; i < sortedRows.length; i++) {
        var rowNum = sortedRows[i];
        if (shouldDebounceRow_(spreadId, sheetId, rowNum, 3000)) continue;

        var rowData = dataToRead[rowNum - sortedRows[0]];
        if (rowData.every(function(c) { return (c === null || String(c).trim() === ''); })) continue;

        var first = (rowData[idxFirst] || '').toString().trim();
        var last = (rowData[idxLast] || '').toString().trim();
        var normDate = normalizeDateValue(rowData[idxDate] || new Date(), new Date().getFullYear());
        var normDob = normalizeDateValue(rowData[idxDOB], new Date().getFullYear());
        
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
            spreadId,
            'Open',
            sheet.getName(),
            ss.getName(),
            srcUuid,
            '' // Placeholder for potential duplicates
        ];
        
        var rowToFormat;
        if (isNewRecord) {
            masterSheet.appendRow(outputRow);
            rowToFormat = masterSheet.getLastRow();
        } else {
            masterSheet.getRange(masterRowFound, 1, 1, outputRow.length).setValues([outputRow]);
            rowToFormat = masterRowFound;
        }

        // Apply formatting and rich text link
        var masterRowRange = masterSheet.getRange(rowToFormat, 1, 1, masterSheet.getLastColumn());
        masterRowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
        if(masterDateColIndex > 0) masterSheet.getRange(rowToFormat, masterDateColIndex).setNumberFormat('MM/dd/yyyy');
        if(masterDobColIndex > 0) masterSheet.getRange(rowToFormat, masterDobColIndex).setNumberFormat('MM/dd/yyyy');
        masterSheet.getRange(rowToFormat, masterLinkColIndex).setRichTextValue(
            SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(sourceRowLink).build()
        );
        
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
    var appendedRange = destSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
    appendedRange.setValues(rowsToAppend);
    appendedRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

    var linkColIndex = destHeaders.indexOf('SHEET LINK TO ROW') + 1;
    var dateColIndex = destHeaders.indexOf('DATE') + 1;
    var dobColIndex = destHeaders.indexOf('DOB') + 1;

    if (dateColIndex > 0) destSheet.getRange(startRow, dateColIndex, rowsToAppend.length, 1).setNumberFormat('MM/dd/yyyy');
    if (dobColIndex > 0) destSheet.getRange(startRow, dobColIndex, rowsToAppend.length, 1).setNumberFormat('MM/dd/yyyy');

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
    // Apply formatting to the whole sheet after clearing
    var fullSheetRange = master.getRange(1, 1, master.getMaxRows(), master.getMaxColumn());
    fullSheetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    master.getRange(2, 1, master.getMaxRows() - 1, 1).setNumberFormat('MM/dd/yyyy'); // Date
    master.getRange(2, 5, master.getMaxRows() - 1, 1).setNumberFormat('MM/dd/yyyy'); // DOB

  } catch (e) {
    Logger.log('Master cleanup failed (pre): ' + e.message);
  }

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

function resetCursorsAndReimportWithoutClearing() {
  resetConsolidationCursors_();
  resetDuplicateGroupLog_();
  consolidateSheetsIncremental(SOURCE_IDS, MASTER_SPREADSHEET_ID, DEST_SHEET_NAME);
  Logger.log('resetCursorsAndReimportWithoutClearing complete.');
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
