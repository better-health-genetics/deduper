
/**
 * This script should be added to EACH of your SOURCE spreadsheets.
 * It connects to the main Master script (as a library) to provide the sidebar UI.
 */

/**
 * onOpen trigger for SOURCE spreadsheets.
 * Creates a simple menu to launch the DeDuper sidebar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BHG DeDuper')
    .addItem('Open Sidebar', 'showDeDuperSidebar')
    .addToUi();
}

/**
 * Shows the sidebar by calling the library function from the Master script.
 */
function showDeDuperSidebar() {
  // 'MasterLibrary' is the Identifier you will set when adding the library.
  const ui = MasterLibrary.getSidebarUiForLibrary();
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * This function acts as a bridge. The sidebar UI calls this function,
 * which in turn calls the actual processing logic in the Master library.
 * This is required because the UI cannot directly call a library function.
 * @param {object} formData The data from the sidebar form.
 * @returns {object} The result from the library's processing function.
 */
function processRecord(formData) {
  // 'MasterLibrary' is the Identifier you will set when adding the library.
  return MasterLibrary.processSourceSheetRecordForLibrary(formData);
}
