# BHG DeDuper & Data Consolidator

## 1. Overview

This project is a robust data management system built for Google Sheets, designed to solve two primary challenges for Better Health Genetics (BHG):

1.  **Data Consolidation**: It automatically aggregates data from numerous source spreadsheets into a single, unified "Master" sheet.
2.  **Manual Data Entry & Duplicate Prevention**: It provides a user-friendly sidebar interface within Google Sheets for adding new records, which performs real-time duplicate checks to ensure data integrity.

The system is composed of a powerful Google Apps Script backend and a modern React-based frontend, providing a seamless and efficient user experience.

---

## 2. Key Features

### Automated Data Consolidation
- **Multi-Source Aggregation**: Pulls data from a configurable list of source Google Sheets.
- **Real-time & Batch Updates**: Consolidates data in near real-time using `onEdit` triggers in source sheets and runs a daily batch process to catch any missed updates.
- **Intelligent Duplicate Marking**: Automatically identifies and flags potential duplicates within the consolidated Master sheet.
- **Comprehensive Logging**: Maintains detailed logs for all operations, errors, and identified duplicates.
- **Admin Controls**: Provides menu items for administrative tasks like full data rebuilds, historical imports, and trigger setup.

### Duplicate Checker Sidebar
- **Modern UI**: A clean, responsive interface built with React and Tailwind CSS that runs directly inside Google Sheets.
- **Efficient Data Entry**: Simple form for adding new records with First Name, Last Name, and Date of Birth.
- **Real-time Duplicate Checking**: Instantly checks for duplicates against the Master sheet upon submission.
- **Smart Record Management**: If a duplicate is found, the system intelligently **updates** the timestamp of the existing record instead of creating a new one. Otherwise, it adds a new record.
- **Live "Duplicate Health" Dashboard**: Displays a color-coded percentage of duplicates found over the last week, giving an at-a-glance view of data quality.
- **Clear User Feedback**: Provides instant visual feedback (success, warning, error messages) to the user after each action.

---

## 3. System Architecture

The entire system is powered by a single Google Apps Script project attached to the **Master Google Sheet**.

-   **Backend (`Code.gs`)**: This is the core engine of the system.
    -   It houses all the data consolidation logic, triggers, and duplicate detection algorithms.
    -   It also serves as the API backend for the frontend sidebar, exposing functions via `google.script.run`.

-   **Frontend (`index.html`)**: This file contains the complete sidebar application.
    -   It's a React application where the JSX is transpiled in the browser by Babel.
    -   It communicates with the `Code.gs` backend to fetch data and submit new records.
    -   It includes a mock server for easy local development and testing of the UI without needing to deploy.

---

## 4. Setup and Deployment

The setup involves configuring a single Google Sheet with the provided Apps Script code.

1.  **Prepare the Master Sheet**:
    *   Create a new Google Sheet. This will be your Master Sheet.
    *   Note its Spreadsheet ID from the URL (`.../spreadsheets/d/SPREADSHEET_ID/edit`).

2.  **Install the Script**:
    *   Open the Script Editor by going to `Extensions` > `Apps Script`.
    *   Create a script file named `Code.gs` and paste the entire contents of the provided `Code.gs` file.
    *   Create an HTML file named `index.html` (`File` > `New` > `HTML file`) and paste the contents of the provided `index.html` file.

3.  **Configure the Script (`Code.gs`)**:
    *   In `Code.gs`, find the configuration variables at the top.
    *   Set `MASTER_SPREADSHEET_ID` to the ID of the sheet you just created.
    *   Update the `SOURCE_IDS` array with the spreadsheet IDs of all the source sheets you want to pull data from.

4.  **First-Time Run & Authorization**:
    *   Save the project.
    *   From the Script Editor, select the `onOpen` function from the dropdown and click **Run**.
    *   This will prompt you to grant the necessary permissions for the script to run. Follow the on-screen instructions to authorize it.

5.  **Initialize the System**:
    *   Go back to your Master Google Sheet and refresh the page. A new menu named **"Data Consolidator"** should appear.
    *   The sidebar should also open automatically.
    *   Use the menu to run the initial setup functions:
        *   `Data Consolidator` > `⚙️ Create Source onEdit Triggers`
        *   `Data Consolidator` > `⏰ Create Daily Consolidation (2am)`
        *   (Optional) `Data Consolidator` > `🔄 Rebuild Master (Full Reset)` to perform an initial full import of all data.

The system is now fully configured and operational.
