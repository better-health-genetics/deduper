/**
 * ===============================================================================================
 *                        BHG DeDuper - Source Sheet Loader Script
 * ===============================================================================================
 *
 * INSTRUCTIONS:
 * 1. Open the script editor for a Source Google Sheet (`Extensions > Apps Script`).
 * 2. Delete any existing code in the `Code.gs` file.
 * 3. Copy and paste THIS ENTIRE SCRIPT into the `Code.gs` file.
 * 4. Add the Master Script as a library:
 *    a. In the left-hand menu, click the `+` icon next to "Libraries".
 *    b. In the "Script ID" field, paste the Script ID of the main "BHG DeDuper" script project.
 *    c. Click "Look up".
 *    d. Choose the latest version.
 *    e. Change the "Identifier" to `BHG_DeDuper`.
 *    f. Click "Add".
 * 5. Save the project.
 * 6. Refresh your Source Sheet. A "DeDuper" menu should now appear.
 *
 * ===============================================================================================
 */

/**
 * Creates a simple menu in the spreadsheet when it's opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('DeDuper')
      .addItem('Show Sidebar', 'showAppSidebar')
      .addSeparator()
      .addItem('Run Diagnostics', 'runDiagnosticsInSourceSheet')
      .addToUi();
}

/**
 * A wrapper function that calls the main library function to show the sidebar.
 */
function showAppSidebar() {
  // `BHG_DeDuper` is the identifier we set when adding the library.
  BHG_DeDuper.showDuplicateCheckerSidebar();
}

/**
 * Calls the new, comprehensive diagnostics function in the master library.
 */
function runDiagnosticsInSourceSheet() {
  try {
    var result = BHG_DeDuper.runDiagnostics();
    
    var htmlOutput = HtmlService.createHtmlOutput(formatDiagnosticsAsHtml(result))
        .setWidth(600)
        .setHeight(450);
        
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Connection Diagnostics Results');
    
  } catch (e) {
    var errorMessage = 
        'An error occurred while trying to call the library. This usually means the library is not attached correctly, the identifier is wrong, or a new version needs to be deployed.\n\n' +
        'Error: ' + e.message;
    SpreadsheetApp.getUi().alert('Library Call Failed', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Formats the diagnostic result object into a clean HTML string for a modal dialog.
 * @param {object} result The result object from runDiagnostics().
 * @return {string} The HTML string.
 */
function formatDiagnosticsAsHtml(result) {
  function escapeHtml(text) {
    if (typeof text !== 'string') text = JSON.stringify(text, null, 2);
    return text
         .replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#039;");
  }

  return `
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
      h2 { color: #444; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
      table { border-collapse: collapse; width: 100%; margin-top: 10px; }
      th, td { text-align: left; padding: 8px; border: 1px solid #ddd; }
      th { background-color: #f2f2f2; width: 150px; }
      td { word-wrap: break-word; word-break: break-all; }
      pre { background-color: #eee; padding: 10px; border-radius: 4px; white-space: pre-wrap; font-family: monospace; }
      .success { color: green; font-weight: bold; }
      .failed { color: red; font-weight: bold; }
    </style>
    <body>
      <h2>Diagnostic Report</h2>
      <table>
        <tr><th>Overall Status</th><td><span class="${result.status === 'OK' ? 'success' : 'failed'}">${escapeHtml(result.status)}</span></td></tr>
        <tr><th>Executing User</th><td>${escapeHtml(result.user)}</td></tr>
        <tr><th>Active Sheet Name</th><td>${escapeHtml(result.activeSheetName)}</td></tr>
        <tr><th>Active Sheet ID</th><td>${escapeHtml(result.activeSheetId)}</td></tr>
      </table>
      
      <h2>Master Sheet Check</h2>
      <table>
        <tr><th>Access Status</th><td><span class="${result.masterAccess === 'Success' ? 'success' : 'failed'}">${escapeHtml(result.masterAccess)}</span></td></tr>
        <tr><th>Details</th><td>${escapeHtml(result.masterAccessMessage)}</td></tr>
      </table>
      
      <h2>getContext() Result</h2>
      <p>This is the critical test. It shows what the backend function sees when called by the UI.</p>
      <pre>${escapeHtml(result.getContextResult)}</pre>

      <h2>Overall Message</h2>
      <pre>${escapeHtml(result.message)}</pre>
    </body>
  `;
}


/**************************************************************************************************
 *                             --- CLIENT-SIDE API BRIDGE ---
 * 
 * The functions below are REQUIRED. They act as a bridge between the client-side HTML UI
 * (which uses `google.script.run`) and the backend library functions. When the sidebar is
 * opened from a Source Sheet, these global functions are exposed to the UI, which then call
 * the actual logic in the `BHG_DeDuper` library.
 **************************************************************************************************/

function getContext() {
  return BHG_DeDuper.getContext();
}

function getSourceSheetHealthData(sourceId) {
  return BHG_DeDuper.getSourceSheetHealthData(sourceId);
}

function addRecordToSourceSheet(formData) {
  return BHG_DeDuper.addRecordToSourceSheet(formData);
}

// These are for the 'OTHER' context but are included for robustness
function getDuplicateHealthData() {
  return BHG_DeDuper.getDuplicateHealthData();
}

function addRecordAndCheckDuplicates(formData) {
  return BHG_DeDuper.addRecordAndCheckDuplicates(formData);
}

