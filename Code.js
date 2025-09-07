
/****************************************************************************************************
 * ___  ____ ___ ___ ____ ____    _  _ ____ ____ _    ___ _  _    ____ ____ _  _ ____ ___ _ ____ ____ 
 * |__] |___  |   |  |___ |__/    |__| |___ |__| |     |  |__|    | __ |___ |\ | |___  |  | |    [__  
 * |__] |___  |   |  |___ |  \    |  | |___ |  | |___  |  |  |    |__] |___ | \| |___  |  | |___ ___] 
 * 
 *                             BETTER HEALTH GENETICS — DATA CONSOLIDATOR
 ****************************************************************************************************/

// =================================================================================================
// MAIN SCRIPT FILE: This file contains the primary onOpen and doGet triggers for the application.
// All other logic has been refactored into separate files for better organization.
// =================================================================================================


/**
 * Auto-add menu when the Master spreadsheet opens.
 * This function now ONLY creates the Admin menu in the Master sheet.
 * Simplified menus for other sheets are handled by a separate stub script in those sheets.
 */
function onOpen() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  if (ssId === MASTER_SPREADSHEET_ID) {
    const ui = SpreadsheetApp.getUi();
    // Show the full admin menu ONLY in the Master Sheet
    ui.createMenu('Data Consolidator')
      .addItem('🔄 Rebuild Master (Full Reset)', 'ui_rebuildMasterFullReset_')
      .addItem('📥 Reimport History (Keep Master)', 'ui_reimportHistoryNoClear_')
      .addSeparator()
      .addItem('Show DeDuper Sidebar', 'showDuplicateCheckerSidebar')
      .addItem('Open Stats Dashboard', 'ui_openDashboard_')
      .addSeparator()
      .addItem('🔑 Set Gemini API Key', 'ui_setApiKey_')
      .addSeparator()
      .addItem('🧭 Seed Source Last Rows', 'ui_initSourceLastRows_')
      .addItem('🧭 Seed Consolidation Cursors', 'ui_initConsolidateLastRows_')
      .addSeparator()
      .addItem('⚙️ Create Source onEdit Triggers', 'ui_createTriggersForSources_')
      .addItem('⏰ Create Daily Consolidation (2am)', 'ui_createDailyConsolidationTrigger_')
      .addSeparator()
      .addItem('📜 Open Log Sheet', 'ui_openLogSheet_')
      .addToUi();
  }
}


/**
 * Serves the web app dashboard. Checks user permission by verifying they can open the Master sheet.
 */
function doGet(e) {
  try {
    // This line acts as a permission check. If the user cannot access the sheet, it will throw an error.
    SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h1>Access Denied</h1><p>You do not have permission to view this page. Please request view access to the Master Spreadsheet.</p>');
  }

  return HtmlService.createHtmlOutput(getDashboardHtml_())
    .setTitle('BHG Duplication Stats Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}