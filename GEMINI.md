# Gemini Code Assistant Guide: BHG DeDuper System

This document provides a technical overview of the BHG DeDuper & Data Consolidator system, tailored for a developer or AI assistant tasked with understanding and modifying the codebase.

## 1. System Architecture

The project is a **monolithic Google Apps Script application** bound to the central "Master" Google Sheet. The backend has been refactored into a modular structure for maintainability.

### 1.1. Key Files

-   **`Config.js`**: **Central Hub for Settings.** Contains all global configuration variables (`SOURCE_IDS`, `MASTER_SPREADSHEET_ID`, sheet names, etc.). This is the first place to look when changing environments or adding new sources.
-   **`Code.js`**: **Main Entry Point.** Contains only the top-level `onOpen` and `doGet` triggers required by Apps Script. It delegates all functional logic to other modules.
-   **`UI.js`**: **User Interface Server.** Manages all functions that present an interface to the user, including `showDuplicateCheckerSidebar`, menu creation functions, and the `getDashboardHtml_` function that serves the full-page web app.
-   **`SidebarAPI.js`**: **Frontend-to-Backend Bridge.** Contains all server-side functions that are directly callable from the frontend UI via `google.script.run`. This file acts as the dedicated API layer for the sidebar.
-   **`Consolidator.js`**: **Core Data Engine.** Houses the primary data processing logic. This includes the `onEditSourceInstallable` trigger, the daily `consolidateSheetsIncremental` batch process, and the `findAndMarkDuplicates` algorithm.
-   **`Helpers.js`**: **Utility Library.** A collection of stateless helper functions used across the other scripts (e.g., `normalizeDateValue`, `findMasterRowByUuid`, `columnToLetter`).
-   **`index.html`**: **The Frontend Application.** A self-contained React app rendered as a Google Sheet sidebar. It communicates with `SidebarAPI.js` via the `google.script.run` bridge.

---

## 2. Core Logic & Patterns

### Context-Awareness

The primary architectural pattern of the UI is context-awareness. The flow is as follows:
1.  The React app in `index.html` mounts and its first action is to call `server.run('getContext')`.
2.  The `getContext()` function in `SidebarAPI.js` inspects the active spreadsheet's ID to determine if it's the Master, a known Source, or Other. For 'Other' sheets, it also checks for the presence of required headers.
3.  It returns a context object (e.g., `{ context: 'MASTER' }` or `{ context: 'SOURCE', sourceId: '...', sourceName: '...' }`).
4.  The React app uses this context object to route to the correct view (AdminDashboard, SourceSheetView, or DefaultView).

### Data Entry Workflows

There are two distinct workflows for adding records, which is critical to understand.

#### A) Source Sheet Workflow
This is the standard flow for data entry clerks working in a designated source sheet.
1.  **UI Action**: User submits the form in the `SourceSheetView`.
2.  **API Call**: `server.run('addRecordToSourceSheet', formData)` is called.
3.  **Backend (`SidebarAPI.js`)**: The `addRecordToSourceSheet` function is simple. It finds the required columns (`FIRST NAME`, `LAST NAME`, `DOB`) and appends a new row **to the active source sheet**.
4.  **Trigger (`Consolidator.js`)**: This action of adding a row fires the `onEditSourceInstallable` trigger for that source sheet.
5.  **Core Logic**: The `onEdit` trigger handles the heavy lifting: it assigns a `MASTER_UUID`, syncs the new record to the `Master` sheet, checks for duplicates, and writes feedback (like potential duplicate links) back to the source sheet row.

#### B) Generic Sheet Workflow
This flow is for ad-hoc entries from any non-standard sheet.
1.  **UI Action**: User submits the form in the `DefaultView`.
2.  **API Call**: `server.run('addRecordAndCheckDuplicates', formData)` is called.
3.  **Backend (`SidebarAPI.js`)**: The `addRecordAndCheckDuplicates` function performs a more complex, direct operation:
    - It finds or creates a `Checker` sheet **locally** (in the user's current spreadsheet) and **centrally** (in the Master spreadsheet).
    - It queries the `Master` sheet to see if a record with the same details already exists.
    - **If a duplicate exists**, it updates the timestamp of the existing record in the `Master` sheet.
    - **If it's a new record**, it appends a new row to the `Master` sheet.
    - It **always** adds a log entry to both the local and central `Checker` sheets, providing a complete audit trail.
4.  **Response**: It returns a detailed status message to the UI.

### Performance Optimization

-   **Fast Lookups**: The script avoids slow `.getValues()` loops for single-record lookups. Instead, `findMasterRecordByDetails_` in `Helpers.js` uses a temporary sheet with a `=QUERY()` formula, delegating the search to Google's highly optimized backend.
-   **Debouncing**: The `onEdit` trigger uses a script cache (`shouldDebounceRow_`) to prevent the same row from being processed multiple times in rapid succession (e.g., from a fast paste).

---

## 3. How to Modify

-   **Change Configuration**: Edit the variables in **`Config.js`**. This is where you'll add new source sheet IDs or change target sheet names.
-   **UI Changes**: All modifications to the sidebar's appearance or behavior should be made to the React components within the `<script type="text/babel">` tag in **`index.html`**.
-   **Add a New API Endpoint**:
    1.  Define the new function in **`SidebarAPI.js`**. This keeps the API layer clean.
    2.  In `index.html`, call it from the frontend using `await server.run('myNewFunction', args)`.
    3.  Add a mock implementation for the new function to the mock server in `index.html` to enable local development.
-   **Modify Core `onEdit` Behavior**: Changes to how data is synced from source sheets to the Master should be made in `onEditSourceInstallable` within **`Consolidator.js`**.
-   **Alter Batch Processing**: To change the nightly consolidation, modify `consolidateSheetsIncremental` in **`Consolidator.js`**.
-   **Add a Helper Function**: If you have a new, reusable utility (e.g., a new data cleaning function), add it to **`Helpers.js`**.