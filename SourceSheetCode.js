
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
