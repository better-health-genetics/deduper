

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
      .addItem('Debug Connection', 'debugLibraryConnection')
      .addToUi();
}

/**
 * A wrapper function that calls the main library function to show the sidebar.
 */
function showAppSidebar() {
  // `BHG_DeDuper` is the identifier we set when adding the library.
  // `showDuplicateCheckerSidebar` is the name of the function in the library's UI.js file.
  BHG_DeDuper.showDuplicateCheckerSidebar();
}

/**
 * Calls the debug function in the master library to diagnose connection/permission issues.
 */
function debugLibraryConnection() {
  SpreadsheetApp.getUi().alert('Running debug check... please wait.');
  try {
    // Call the debug function from the library
    var result = BHG_DeDuper.debugConnection();
    
    // Format the result object into a readable string for the alert box
    var message = 
        'Connection Status: ' + result.status + '\n\n' +
        'User: ' + result.user + '\n' +
        'Master Sheet Access: ' + result.masterAccess + '\n' +
        'Active Sheet ID: ' + result.activeSheetId + '\n' +
        'Active Sheet Name: ' + result.activeSheetName + '\n\n' +
        'Message: ' + result.message;
        
    SpreadsheetApp.getUi().alert('Connection Debug Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (e) {
    var errorMessage = 
        'An error occurred while trying to call the library. This usually means the library is not attached correctly, the identifier is wrong, or a new version needs to be deployed.\n\n' +
        'Error: ' + e.message;
    SpreadsheetApp.getUi().alert('Library Call Failed', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}