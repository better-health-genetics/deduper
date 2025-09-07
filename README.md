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
The sidebar's UI and functionality intelligently adapt to the user's current location within Google Sheets.

-   **Admin Dashboard View (in Master Sheet)**:
    -   When opened in the Master Sheet, it displays a comprehensive dashboard of duplication statistics for all source sheets.
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
- A full-page, graphical dashboard accessible via a shareable URL.
- Presents duplication statistics as an interactive bar chart, offering a clean, high-level view for stakeholders.
- Access is restricted to users who have at least view permissions on the Master Sheet.

---

## 3. System Architecture

The entire system is powered by a Google Apps Script project attached to the **Master Google Sheet**. The backend has been refactored into a modular, maintainable structure.

-   **Backend (Multiple `.js` files)**:
    -   **`Config.js`**: Centralized configuration for all spreadsheet IDs and settings.
    -   **`Code.js`**: Main entry point containing `onOpen` and `doGet` triggers.
    -   **`UI.js`**: Manages all UI-related actions like showing the sidebar and serving the web app.
    -   **`SidebarAPI.js`**: The API layer containing all functions called by the frontend.
    -   **`Consolidator.js`**: The core data processing engine, including the `onEdit` trigger and batch consolidation logic.
    -   **`Helpers.js`**: A collection of utility functions used across the backend.

-   **Frontend (`index.html`)**:
    -   A complete, self-contained React application that runs in the Google Sheets sidebar.
    -   Communicates with the backend API functions via `google.script.run`.
    -   Includes a mock server for easy local development and testing of the UI.

---

## 4. Setup and Deployment

1.  **Prepare the Master Sheet**:
    *   Create a new Google Sheet. This will be your Master Sheet.
    *   Note its Spreadsheet ID from the URL (`.../spreadsheets/d/SPREADSHEET_ID/edit`).

2.  **Install the Script**:
    *   Open the Script Editor by going to `Extensions` > `Apps Script`.
    *   Delete any existing files.
    *   Create a new script file for **each** of the backend files (`Code.js`, `Config.js`, `UI.js`, etc.), making sure the filenames match exactly. Copy and paste the contents into the corresponding files.
    *   Create an HTML file named `index.html` (`File` > `New` > `HTML file`) and paste its contents.

3.  **Configure the Script (`Config.js`)**:
    *   Open the `Config.js` file in the script editor.
    *   Set `MASTER_SPREADSHEET_ID` to the ID of the sheet you just created.
    *   Update the `SOURCE_IDS` array with the spreadsheet IDs of all the source sheets you want to pull data from.

4.  **First-Time Run & Authorization**:
    *   Save the project.
    *   From the Script Editor, select the `onOpen` function from the dropdown and click **Run**.
    *   This will prompt you to grant the necessary permissions. Follow the on-screen instructions to authorize it.

5.  **Initialize the System**:
    *   Go back to your Master Google Sheet and refresh the page. A new menu named **"Data Consolidator"** should appear, and the sidebar should open automatically.
    *   Use the menu to run the initial setup functions:
        *   `Data Consolidator` > `⚙️ Create Source onEdit Triggers`
        *   `Data Consolidator` > `⏰ Create Daily Consolidation (2am)`
        *   (Optional) `Data Consolidator` > `🔄 Rebuild Master (Full Reset)` to perform an initial full import of all data.

The system is now fully configured and operational.