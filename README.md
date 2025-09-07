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
- **Intelligent Duplicate Marking**: Automatically identifies and flags potential duplicates within the consolidated Master sheet.
- **Comprehensive Logging**: Maintains detailed logs for all operations, errors, and identified duplicates.
- **Admin Controls**: Provides menu items for administrative tasks like full data rebuilds, historical imports, and trigger setup.

### Context-Aware Sidebar
The sidebar's UI intelligently adapts to the user's current location.

-   **Admin Dashboard View (in Master Sheet)**:
    -   When opened in the Master Sheet, it displays a comprehensive dashboard of duplication statistics for all source sheets.
    -   **New! AI Health Summary**: Features a "Generate Analysis" button that uses the Google Gemini API to provide a natural language summary of the data quality, highlighting problem areas and successes.
    -   Features dynamic date-range filtering (WTD, MTD, YTD, ALL) to analyze data quality over different periods.
    -   Provides direct links to each source sheet for quick navigation.

-   **Source Sheet View (in a configured Source Sheet)**:
    -   Provides a simple form for adding new records (`First Name`, `Last Name`, `DOB`).
    -   Displays "Duplicate Health" statistics relevant **only to that specific source sheet**.
    -   Submitting the form adds the new record **directly to the current sheet**, which is then automatically synced to the Master sheet by the `onEdit` trigger.

-   **Generic View (in any other sheet)**:
    -   Activates a data entry form only if the sheet contains `FIRST NAME`, `LAST NAME`, and `DOB` columns.
    -   Displays **overall** duplicate health statistics from the entire Master sheet.
    -   Submitting a record logs the entry to a local `Checker` sheet and a central `Checker` sheet in the Master file, then writes the data directly to the Master sheet.


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
    -   This "stub" script's only job is to **include the Master library** and use it to create a menu and show the sidebar. This makes the connection explicit and easy to manage.

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
    *   Run the initial setup commands: `⚙️ Create Source onEdit Triggers` and `⏰ Create Daily Consolidation (2am)`.

### Part B: Setting up a Source Sheet

**Repeat these steps for EACH of your source sheets.**

1.  **Open the Source Sheet** and go to its Script Editor (`Extensions > Apps Script`).
2.  You should see a file named `Code.gs`. Delete any code inside it.
3.  Copy the entire contents of the `SourceSheetCode.js` file and paste it into `Code.gs`.
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

If the sidebar in a Source Sheet shows a "Could not connect to the backend" error, use the built-in debugger:
1. In the Source Sheet, go to the `DeDuper` menu.
2. Click `Debug Connection`.
3. An alert box will appear with diagnostic information. The most important line is **Master Sheet Access**.
4. If it says `FAILED`, it means the current user does not have at least **view permission** for the Master Google Sheet. Grant them view access and try again.
5. If the entire debug action fails with an error, it's likely the library was not added correctly in the Source Sheet's script project. Double-check that the library is added with the identifier `BHG_DeDuper` and that you are using the latest version of the library.