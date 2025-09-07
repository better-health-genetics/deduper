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

This model is more robust than a single monolithic script because it makes the connection between the source sheets and the main logic explicit and easier to manage.

### 1.1. Key Files (In Master Script Project)

-   **`Config.js`**: **Central Hub for Settings.** Contains all global configuration variables (`SOURCE_IDS`, `MASTER_SPREADSHEET_ID`, `WEB_APP_URL`, etc.).
-   **`Code.js`**: **Main Entry Point.** Contains the `onOpen` for the admin menu and the `doGet` for the web dashboard.
-   **`UI.js`**: **User Interface Server.** Contains `showDuplicateCheckerSidebar()` (which is exposed to the library) and the HTML generation for the dashboard.
-   **`SidebarAPI.js`**: **Frontend-to-Backend Bridge.** Contains all server-side functions directly called by the React frontend via `google.script.run`.
-   **`Consolidator.js`**: **Core Data Engine.** Houses the `onEditSourceInstallable` trigger, batch processing, and duplicate finding logic.
-   **`Helpers.js`**: **Utility Library.** A collection of stateless helper functions.
-   **`index.html`**: **The Frontend Application.** A self-contained React app.
-   **`SourceSheetCode.js`**: **Template for Consumers.** A template file containing the code to be placed in each source sheet's script project.

---

## 2. Core Logic & Patterns

### Context-Awareness

Context is still determined when the sidebar loads, but the menu creation is now decentralized.
1.  A user in a Source Sheet clicks the menu item created by the local `SourceSheetCode.js` script.
2.  This calls the library function `BHG_DeDuper.showDuplicateCheckerSidebar()`.
3.  The sidebar UI (`index.html`) loads. Its first action is to call `server.run('getContext')`.
4.  The `getContext()` function (in the library's `SidebarAPI.js`) inspects the active spreadsheet's ID to determine if it's the Master, a known Source, or Other.
5.  It returns a context object (e.g., `{ context: 'SOURCE', sourceId: '...' }`).
6.  The React app uses this context object to render the correct view (AdminDashboard, SourceSheetView, etc.).

### Data Entry Workflows

The actual data processing workflows remain the same as described previously, as all the logic still resides in the central library. The only change is how the user interface is initiated.

-   **Source Sheet Workflow**: The `onEditSourceInstallable` trigger is still created and managed by the Master Script admin, but the sidebar that initiates new entries is now opened via the local script.
-   **Generic Sheet Workflow**: If a user adds the library to a generic sheet, the `getContext` call will identify it as 'OTHER' or 'OTHER_NO_HEADERS', and the sidebar will function in "Checker Mode" as designed.

---

## 3. How to Modify

-   **Change Configuration**: Edit the variables in **`Config.js`** in the Master Script project.
-   **UI Changes**: All modifications to the sidebar's appearance or behavior should be made in **`index.html`** in the Master Script project.
-   **Add a New API Endpoint**:
    1.  Define the new function in **`SidebarAPI.js`** in the Master Script project.
    2.  Call it from `index.html` using `await server.run('myNewFunction', args)`.
-   **Modify `onEdit` Behavior**: Changes to how data is synced from source sheets to the Master should be made in `onEditSourceInstallable` within **`Consolidator.js`** in the Master Script project.
-   **Expose a New Library Function**: If you need to call a new top-level function from the source sheet stubs, ensure it is defined globally (e.g., in `UI.js` or `Code.js`) in the Master Script project.
-   **Deploying Changes**: After making changes to the Master Script, you must create a **new version** of the library deployment (`Deploy > Manage deployments > Select Library deployment > Edit > New version`). Source sheets will then automatically use the updated version.