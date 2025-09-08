# BHG DeDuper & Data Consolidator

## 1. Overview

This project is a robust data management system built for Google Sheets, designed to solve two primary challenges for Better Health Genetics (BHG):

1.  **Data Consolidation**: It automatically aggregates data from numerous source spreadsheets into a single, unified "Master" sheet.
2.  **Manual Data Entry & Duplicate Prevention**: It provides a user-friendly, context-aware sidebar for adding new records. The sidebar performs real-time duplicate checks and behaves differently depending on the sheet it's opened in, ensuring a safe and intuitive workflow for all users.

The system is composed of a powerful, modular Google Apps Script backend and a modern React-based frontend, providing a seamless and efficient user experience.

---

## 2. Key Features

### Automated Data Consolidation
- **Multi-Source Aggregation**: Pulls data from a configurable list of source Google Sheets.
- **Real-time & Batch Updates**: Consolidates data in near real-time using `onEdit` triggers in source sheets and runs a daily batch process to catch any missed updates.
- **Intelligent Duplicate Marking**: Automatically identifies and flags potential duplicates within the consolidated Master sheet using a combination of First Name, Last Name, and DOB.
- **Comprehensive Logging**: Maintains detailed logs for all operations, errors, and identified duplicates.
- **Admin Controls**: Provides menu items in the Master Sheet for administrative tasks like full data rebuilds, historical imports, and trigger setup.

### Context-Aware Sidebar
The sidebar's UI intelligently adapts to the user's current location.

-   **Admin Dashboard View (in Master Sheet)**:
    -   When opened in the Master Sheet, it displays a comprehensive dashboard of duplication statistics for all source sheets.
    -   **AI Health Summary**: Features a "Generate Analysis" button that uses the Google Gemini API to provide a natural language summary of the data quality, highlighting problem areas and successes.
    -   Features dynamic date-range filtering (WTD, MTD, YTD, ALL) to analyze data quality over different periods.
    -   Provides direct links to each source sheet for quick navigation.

-   **Source Sheet View (in a configured Source Sheet)**:
    -   Provides a simple form for adding new records (`First Name`, `Last Name`, `DOB`).
    -   Displays "Duplicate Health" statistics relevant **only to that specific source sheet**.
    -   Submitting the form adds the new record **directly to the current sheet**, which is then automatically synced to the Master sheet by the `onEdit` trigger.

-   **Generic View (in any other sheet)**:
    -   Activates a data entry form only if the sheet contains `FIRST NAME`, `LAST NAME`, and `DOB` columns.
    -   Displays **overall** duplicate health statistics from the entire Master sheet.
    -   Submitting a record logs the entry to a local `Checker` sheet and a central `Checker` sheet in the Master file, then writes the data directly to the Master sheet, checking for duplicates in the process.


### Standalone Web Dashboard
- A full-page, data-rich dashboard accessible via a shareable URL.
- Presents duplication statistics with high-level summary cards, a bar chart for visual comparison, and a detailed, sortable table for granular analysis.
- Access is restricted to users who have at least view permissions on the Master Sheet.

---

## 3. System Architecture

The system uses a **library-based architecture** for scalability and manageability.

-   **Master Script (The Library)**:
    -   A central Google Apps Script project bound to the Master Sheet. It contains all the core logic, the UI (`index.html`), and all the backend processing functions.
    -   This project is deployed as a **library**, allowing other scripts to call its functions.

-   **Source Sheet Scripts (The Consumers)**:
    -   Each source spreadsheet has its own, very small, container-bound script.
    -   This "stub" script's only job is to **include the Master library** and use it to create a menu and expose the necessary API functions for the UI. This is critical because `google.script.run` (used by the sidebar) can only call functions in the local script project.

---

## 4. Setup and Deployment

This is a two-part process: first setting up the central library, then configuring each source sheet to use it.

### Part A: Setting up the Master Library

1.  **Prepare the Master Sheet**:
    *   Create a new Google Sheet. This will be your Master Sheet.
    *   Note its Spreadsheet ID from the URL (`.../spreadsheets/d/SPREADSHEET_ID/edit`).

2.  **Install the Script**:
    *   Open the Script Editor in the Master sheet (`Extensions` > `Apps Script`).
    *   Delete any existing files.
    *   Create a script file for **each** backend file (`Code.js`, `Config.js`, etc.) and an HTML file for `index.html`. Copy and paste the contents into the corresponding files.

3.  **Configure the Script (`Config.js`)**:
    *   Open the `Config.js` file.
    *   Set `MASTER_SPREADSHEET_ID` to the ID of your Master sheet.
    *   Update `SOURCE_IDS` with the spreadsheet IDs of all your source sheets.

4.  **Deploy the Web App & Configure URL**:
    *   In the Script Editor, click **Deploy > New deployment**.
    *   Click the **gear icon** and select **Web app**. Configure it as `Execute as: User accessing the web app` and `Who has access: Anyone with Google account`.
    *   Click **Deploy**. Copy the **Web app URL**.
    *   **CRITICAL**: Paste this URL into the `WEB_APP_URL` variable in `Config.js`.

5.  **Deploy as a Library**:
    *   Click **Deploy > New deployment** again.
    *   Click the **gear icon** and select **Library**.
    *   Enter a description (e.g., "BHG DeDuper Core").
    *   Click **Deploy**.
    *   Copy the **Script ID** provided. You will need this for Part B.

6.  **Authorize and Initialize**:
    *   In the Script Editor, select the `onOpen` function and click **Run**. Authorize the script when prompted.
    *   Refresh your Master Sheet. The **"Data Consolidator"** menu should appear.
    *   Go to `Data Consolidator` > `🔑 Set Gemini API Key` and enter your key.
    *   Run the initial setup commands from the menu: `⚙️ Create Source onEdit Triggers` and `⏰ Create Daily Consolidation (2am)`.

### Part B: Setting up a Source Sheet

**Repeat these steps for EACH of your source sheets.**

1.  **Open the Source Sheet** and go to its Script Editor (`Extensions > Apps Script`).
2.  You should see a file named `Code.gs`. Delete any code inside it.
3.  Copy the entire script below and paste it into `Code.gs`.

    ```javascript
    /**
     * ===============================================================================================
     *                        BHG DeDuper - Source Sheet Loader Script
     * ===============================================================================================
     */

    function onOpen() {
      SpreadsheetApp.getUi()
          .createMenu('DeDuper')
          .addItem('Show Sidebar', 'showAppSidebar')
          .addSeparator()
          .addItem('Run Diagnostics', 'runDiagnosticsInSourceSheet')
          .addToUi();
    }

    function showAppSidebar() {
      BHG_DeDuper.showDuplicateCheckerSidebar();
    }
    
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
     * (which uses `google.script.run`) and the backend library functions.
     **************************************************************************************************/

    function getContext() { return BHG_DeDuper.getContext(); }
    function getSourceSheetHealthData(sourceId) { return BHG_DeDuper.getSourceSheetHealthData(sourceId); }
    function addRecordToSourceSheet(formData) { return BHG_DeDuper.addRecordToSourceSheet(formData); }
    function getDuplicateHealthData() { return BHG_DeDuper.getDuplicateHealthData(); }
    function addRecordAndCheckDuplicates(formData) { return BHG_DeDuper.addRecordAndCheckDuplicates(formData); }
    function getAdminDashboardData(startDate, endDate) { return BHG_DeDuper.getAdminDashboardData(startDate, endDate); }
    function getGeminiHealthSummary(stats) { return BHG_DeDuper.getGeminiHealthSummary(stats); }
    ```

4.  **Add the Library**:
    *   In the left-hand menu, click the **`+` icon** next to "Libraries".
    *   Paste the **Script ID** you copied from Part A, Step 5.
    *   Click **Look up**.
    *   Ensure the latest version is selected.
    *   **IMPORTANT**: Change the "Identifier" to `BHG_DeDuper` (this must match the code).
    *   Click **Add**.
5.  **Save and Authorize**:
    *   Save the script project.
    *   From the function dropdown, select `onOpen` and click **Run**. Authorize the script when prompted.
6.  **Done!**: Refresh your source sheet. The simple "DeDuper" menu should now appear, and the sidebar will work correctly, powered by the central library.

### 4.1. Troubleshooting

If the sidebar in a Source Sheet shows a "Could not connect to the backend" or a similar error, use the built-in diagnostics tool:
1. In the Source Sheet, go to the `DeDuper` menu.
2. Click `Run Diagnostics`.
3. A modal dialog will appear with detailed information. Pay close attention to two sections:
    - **Master Sheet Check**: If the `Access Status` is `FAILED`, it means the current user does not have at least **view permission** for the Master Google Sheet. Grant them view access and try again.
    - **getContext() Result**: This shows the exact output or error from the core function that determines how the sidebar should behave. If this section shows an error, it indicates a problem with the script's ability to understand the current sheet environment. Please provide this output to the developer.
4. If the entire diagnostic action fails with an error, it's likely the library was not added correctly in the Source Sheet's script project. Double-check that the library is added with the identifier `BHG_DeDuper` and that you are using the latest version of the library.