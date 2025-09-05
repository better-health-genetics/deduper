# Gemini Code Assistant Guide: BHG DeDuper System

This document provides a technical overview of the BHG DeDuper & Data Consolidator system, tailored for a developer or AI assistant tasked with understanding and modifying the codebase.

## 1. System Architecture

The project is a **monolithic Google Apps Script application** bound to the central "Master" Google Sheet. It integrates a backend for data processing with a service that delivers a frontend UI.

### 1.1. Key Files

-   **`Code.gs`**: The single source of truth for all backend logic. It contains configuration variables (`SOURCE_IDS`, `MASTER_SPREADSHEET_ID`, etc.) at the top. All server-side functions callable from the UI are located here.
-   **`index.html`**: The complete frontend application. It is a React app rendered as a Google Sheet sidebar. The JSX code is located within a `<script type="text/babel">` tag and is transpiled in-browser. It communicates with `Code.gs` via the `google.script.run` bridge.

### 1.2. Backend (`Code.gs`)

The backend has two primary responsibilities:

1.  **Data Consolidation**:
    -   **Triggers**: Uses time-based (daily) and `onEdit` triggers to pull data incrementally from source sheets defined in `SOURCE_IDS`.
    -   **Logic**: The core function is `consolidateSheetsIncremental`, which reads data from source sheets since its last run, normalizes it, and appends it to the `Master` sheet.
    -   **Duplicate Detection**: After data ingestion, `findAndMarkDuplicates` is run. It groups records by a composite key (First Name, Last Name, DOB) and flags groups with more than one entry, writing links to the potential duplicates in the `POTENTIAL_DUPLICATES` column.

2.  **Sidebar API**:
    -   Exposes functions to the frontend via `google.script.run`.
    -   `getDuplicateHealthData()`: Calculates statistics by reading the `Checker` sheet.
    -   `addRecordAndCheckDuplicates()`: The primary entry point for the UI. It handles the core logic for manual record addition.

### 1.3. Frontend (`index.html`)

-   **Framework**: React with Hooks (`useState`, `useEffect`, `useCallback`).
-   **Styling**: Tailwind CSS loaded via CDN.
-   **Backend Communication**: All calls to `Code.gs` are proxied through the `server` object. This object intelligently switches between `google.script.run` (in the live environment) and a built-in mock server (for local browser development).
-   **State Management**: Component-level state is used to manage UI state (loading spinners, form data, messages).

---

## 2. Core Logic & Patterns

### Adding / Checking a Record (`addRecordAndCheckDuplicates`)

This function follows a specific, important pattern: **Update-then-Log**.

1.  **Input**: Receives `formData` object `{firstName, lastName, dob}`.
2.  **Query**: It calls a helper, `findMasterRecordByDetails_`, which performs a fast, non-iterative lookup. It writes a `=QUERY(...)` formula to a hidden helper sheet to find a matching record's `MASTER_UUID` in the `Master` sheet.
3.  **Update or Create**:
    -   **If a match is found (Duplicate)**: It **updates** the `DATE` field of the *existing row* in the `Master` sheet. It does **not** create a new record. The status message reflects this update.
    -   **If no match is found (New Record)**: It **appends** a new row to the `Master` sheet, generating a new `MASTER_UUID`.
4.  **Log**: It **always** appends a new row to the `Checker` sheet. This sheet serves as an immutable transaction log of every submission via the sidebar.
5.  **Response**: It returns a result object to the UI. The `success` property is used to control the message type (e.g., `success: false` for a duplicate finding triggers a `Warning` message).

### Performance Optimization

-   The script avoids slow, iterative `.getValues()` loops on large datasets for lookups.
-   The use of `=QUERY(IMPORTRANGE(...))` or `=QUERY(...)` on a temporary helper sheet delegates the search operation to Google's highly optimized backend, which is significantly faster than Apps Script loops and avoids execution time limits.

### Idempotency and Data Integrity

-   **`MASTER_UUID`**: A universally unique identifier is assigned to every record upon its first entry into the `Master` sheet. This UUID is the canonical identifier for a record and is used for reliable updates.
-   **Composite Keys**: For initial duplicate checks, a composite key of `UPPER(FIRST NAME)|UPPER(LAST NAME)|YYYY-MM-DD(DOB)` is used to group potential matches.

---

## 3. How to Modify

-   **UI Changes**: All modifications to the sidebar's appearance or behavior should be made to the React components within the `<script type="text/babel">` tag in `index.html`.
-   **Backend Logic Changes**: Modify the corresponding functions in `Code.gs`. For example, to change how duplicate health is calculated, edit `getDuplicateHealthData`.
-   **Adding a New API Endpoint**:
    1.  Define a new function in `Code.gs` (e.g., `function myNewFunction(args) { ... }`).
    2.  In `index.html`, call it from the frontend using `await server.run('myNewFunction', args)`.
    3.  (Optional) Add a mock implementation for the new function to the mock server in `index.html` to enable local development.
