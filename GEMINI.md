# Gemini Code Assistant Guide: BHG DeDuper System

This document provides a technical overview of the BHG DeDuper & Data Consolidator system, tailored for a developer or AI assistant tasked with understanding and modifying the codebase.

## 1. System Architecture

The project utilizes a **library-based architecture** for scalability and maintainability.

-   **Master Script (The Library)**:
    -   This is the central Apps Script project, container-bound to the Master Google Sheet.
    -   It contains all the core logic, backend functions (`SidebarAPI.js`, `Consolidator.js`, etc.), and the frontend UI (`index.html`).
    -   It is deployed as a **library**, making its functions available to other scripts.
    -   Its own `onOpen` function is responsible **only** for creating the administrative menu within the Master Sheet itself.

-   **Source Sheet Scripts (The Consumers)**:
    -   Each source spreadsheet has its own simple, container-bound "stub" script (`SourceSheetCode.js`).
    -   The primary responsibilities of this stub script are:
        1.  To include the Master Script as a library (identified as `BHG_DeDuper`).
        2.  To create a simple "DeDuper" menu using its own `onOpen` trigger.
        3.  To call the library's `showDuplicateCheckerSidebar()` function when the menu item is clicked.
        4.  **Crucially**, to provide "bridge" or "wrapper" functions for every backend API endpoint that the UI needs to call. This is necessary because `google.script.run` can only see global functions within the local script project.

### 1.1. Key Files (In Master Script Project)

-   **`Config.js`**: **Central Hub for Settings.** Contains all global configuration variables (`SOURCE_IDS`, `MASTER_SPREADSHEET_ID`, `WEB_APP_URL`, etc.).
-   **`Code.js`**: **Main Entry Point.** Contains the `onOpen` for the admin menu and the `doGet` for the web dashboard.
-   **`UI.js`**: **User Interface Server.** Contains `showDuplicateCheckerSidebar()` (which is exposed to the library), the HTML generation for the standalone web dashboard (`getDashboardHtml_`), and wrapper functions for the admin menu items.
-   **`SidebarAPI.js`**: **Frontend-to-Backend Bridge.** Contains all server-side functions directly called by the React frontend via `google.script.run`. This is the core API layer for the UI.
-   **`Consolidator.js`**: **Core Data Engine.** Houses the `onEditSourceInstallable` trigger logic, the daily batch processing (`consolidateSheetsIncremental`), and the duplicate finding logic (`findAndMarkDuplicates`).
-   **`Helpers.js`**: **Utility Library.** A collection of stateless helper functions for tasks like date normalization, key generation, sheet manipulation, and API key retrieval.
-   **`index.html`**: **The Frontend Application.** A self-contained React/JSX application that renders the sidebar UI.
-   **`SourceSheetCode.js`**: **Template for Consumers.** A template file containing the code to be placed in each source sheet's script project. This file *must* contain the API bridge functions.

---

## 2. Core Logic & Patterns

### Context-Awareness & The Client-Side Bridge

This is the most critical architectural pattern to understand. The client-side UI (`index.html`) cannot directly call functions in the Master Script library because it is running in the context of a Source Sheet. It can only call global functions in the local script project (`SourceSheetCode.js`).

The flow works as follows:

1.  A user in a Source Sheet clicks the menu item, which calls `showAppSidebar()` in the local script.
2.  `showAppSidebar()` calls the library function `BHG_DeDuper.showDuplicateCheckerSidebar()`.
3.  The sidebar UI (`index.html`) loads. Its first action is to call `server.run('getContext')`.
4.  The `google.script.run` object looks for a global function named `getContext` in the **local script** (`SourceSheetCode.js`).
5.  The wrapper function `getContext()` in `SourceSheetCode.js` is found. It simply calls `BHG_DeDuper.getContext()`, passing the request to the main library.
6.  The `getContext()` function in the library's `SidebarAPI.js` runs its logic (checking spreadsheet IDs, etc.) and returns the context object.
7.  The result is passed back through the bridge to the client, which uses it to render the correct view (AdminDashboard, SourceSheetView, etc.).

This bridge pattern is essential for any function that the UI needs to call.

### Data Entry Workflows

-   **Source Sheet Workflow**: The user adds a record via the sidebar. `addRecordToSourceSheet` is called, which writes the data directly to the active source sheet. The `onEditSourceInstallable` trigger (installed by an admin from the Master Sheet) then picks up this change and syncs it to the Master sheet.
-   **Generic Sheet Workflow**: The user adds a record. `addRecordAndCheckDuplicates` is called. This function communicates directly with the Master sheet to check for duplicates, update or create a record, and logs the entry in a "Checker" tab on both the local and Master sheets.

---

## 3. How to Modify

-   **Change Configuration**: Edit the variables in **`Config.js`** in the Master Script project.
-   **UI Changes**: All modifications to the sidebar's appearance or behavior should be made in **`index.html`** in the Master Script project.
-   **Add a New API Endpoint for the UI**:
    1.  **Define the core logic**: Add the new function to **`SidebarAPI.js`** in the Master Script project. This function will contain the actual business logic.
    2.  **Expose the function via the bridge**: Add a corresponding wrapper function to the **`SourceSheetCode.js`** template file (e.g., `function myNewFunction(args) { return BHG_DeDuper.myNewFunction(args); }`).
    3.  **Update consumer scripts**: **This is a critical step.** You must manually copy the updated `SourceSheetCode.js` content into the script project of **every single source sheet** that uses the library.
    4.  **Call from the UI**: You can now call the new endpoint from `index.html` using `await server.run('myNewFunction', args)`.
-   **Modify `onEdit` Behavior**: Changes to how data is synced from source sheets to the Master should be made in `onEditSourceInstallable` within **`Consolidator.js`** in the Master Script project.
-   **Deploying Changes**: After making changes to the Master Script, you must create a **new version** of the library deployment (`Deploy > Manage deployments > Select Library deployment > Edit > New version`). Source sheets will then automatically use the updated version, but this does NOT update the `SourceSheetCode.js` file in those sheets; that must be done manually if the API bridge is changed.