# RFA VBA Process Documentation

**Source workbook:** `Testing/Macro.xlsm`  
**Export folder:** `RFA/exported_vba`  
**Analysis method:** Static inspection only; the macros were not run in Excel.

**Repository safety:** Plaintext password values retained inside the ignored workbook were replaced with `<REDACTED>` in the tracked VBA export before commit. VBA project protection remains enabled in the V6.0 workbook.

The workbook is an Employee Tracking System (ETS) that produces Return From Absence (RFA) reports. Its executable VBA self-references use `ThisWorkbook` rather than a version-specific filename, so future file renames do not require code changes.

## Purpose

The VBA application loads two employee roster cycles, lets an operator select a date, position, and shift, identifies employees relevant to that selection, builds absence-history remarks, writes the report to the workbook, and exports an RFA PDF to a network folder.

The application is built primarily around the `MainWindow` UserForm. Excel is hidden during normal operation, so the form acts as the user-facing application.

## End-to-End Process

```mermaid
flowchart TD
    Open["Open macro-enabled workbook"] --> WOpen["Workbook_Open event"]
    WOpen --> HideExcel["Hide Excel window and show MainWindow"]
    HideExcel --> Init["Initialize form controls and ListViews"]
    Init --> Discover["Scan 03 TG Roster cycle folders"]
    Discover --> Paths["Resolve the two highest consecutive cycles"]
    Paths --> Load1["Validate First Cycle Current Roster"]
    Paths --> Load2["Validate Second Cycle Current Roster"]
    Load1 --> Dates["Build 28-day date-to-cycle map"]
    Load2 --> Dates
    Dates --> Select["Operator selects date, position, and shift"]
    Select --> Available["Calculate available shift values"]
    Select --> Track["Track Now"]
    Track --> Cycle{"Selected date cycle?"}
    Cycle -->|First Cycle| First["Build/filter First Cycle candidates"]
    Cycle -->|Second Cycle| Second["Build/filter Second Cycle candidates"]
    First --> Review["Display employee list and 28-day schedule"]
    Second --> Review
    Review --> Search["Optional employee search and inspection"]
    Review --> Print["Print / create RFA report"]
    Search --> Print
    Print --> History["Scan prior statuses and build absence remarks"]
    History --> Sheet["Write report rows to Sheet1"]
    Sheet --> PDF["Export Sheet1 as RFA Report PDF"]
    PDF --> Close["Save/close workbook and quit Excel"]
```

## Startup and Data Loading

The workbook class module contains the startup events and is exported as `ThisWorkbook.cls` for manual copying into the destination workbook object.

1. `Workbook_Open` hides Excel, disables screen updates and alerts, hides the current workbook window through `ThisWorkbook.Windows(1)`, and displays `MainWindow`. An unexpected startup error restores Excel and reports the captured cause.
2. `UserForm_Initialize` configures the visible and hidden ListView controls, disables tracking and report output, and calls `LoadStartupDataWithRetry`.
3. `TryResolveLatestRosterPaths` scans `\\mcp.com\dept$\FP&A\RFA\03 TG Roster` for directories named `CYCLE <positive integer>`, ignores unrelated names, and selects the two highest cycle numbers.
4. The resolver requires the selected cycle numbers to be consecutive and constructs each filename as `CYCLE <n> ROSTER SUMMARY (Dealer, Pit Supervisor and Pit Manager).xlsx` inside its matching folder. A missing newest file is an error; the loader does not silently fall back to an older cycle.
5. `TryLoadRoster` opens the older cycle as First Cycle and the newest as Second Cycle, both read-only. Each workbook must contain a `Current Roster` worksheet, a region of at least 3 rows by 16 columns, a valid cycle start date in C2, and at least one employee beginning on row 3. The newer C2 date must be exactly 14 days after the older date.
6. Both roster ranges are read into arrays before any visible or hidden form data is cleared. Only after both sources pass does `PopulateRosterControls` replace the form data and build the 28-day, two-cycle date map.
7. The loader restores Excel application flags and enables tracking and the inactivity timer only after a complete successful load.

```mermaid
flowchart LR
    Root["03 TG Roster"] --> Scan["Parse CYCLE number folders"]
    Scan --> Rank["Select two highest numbers"]
    Rank --> Sequence{"Consecutive?"}
    Sequence -->|Yes| R1["Older: First-cycle Current Roster"]
    Sequence -->|Yes| R2["Newest: Second-cycle Current Roster"]
    Sequence -->|No| Choice{"Retry or Cancel?"}
    R1 --> Validate{"Both rosters valid?"}
    R2 --> Validate
    Validate -->|Yes| Memory["Commit arrays to hidden form ListViews"]
    Validate -->|No| Choice
    Choice -->|Retry| Root
    Choice -->|Cancel| Disabled["Reveal Excel; leave tracking disabled; keep Refresh available"]
    Memory --> Map["28-day date and cycle map"]
```

Validation is all-or-nothing: an inaccessible root, missing or nonconsecutive cycle, missing roster, or invalid source never clears previously loaded form data. Error dialogs identify the discovery or roster failure, show the full failing path, describe the Excel error when one is available, and state the corrective action. Retry scans the folder again, so a newly published cycle is picked up automatically. Cancel closes any partially opened source workbook, restores Excel visibility and application state, stops the timer, disables tracking, and leaves Refresh enabled so the operator can retry later.

## Selection and Tracking

The operator works with three main selectors:

- **Date:** mapped to a roster column and either the first or second cycle.
- **Position:** All, Dealer, Pit Supervisor, or Pit Manager.
- **Shift:** All, Morning, Day, Late Day, Night, or Late Night.

Changing any selector stops the inactivity timer, recalculates available shifts through `CallShiftStarts`, and restarts the timer. The available-shift logic delegates to `LookForRecordFirstForShifting` or `LookForRecordSecondForShifting` according to the selected cycle.

`cmdTrackNow_Click` is the main tracking entry point. It:

1. Resolves the selected date to a roster column and cycle.
2. Clears previous results and display state.
3. Calls `FirstCycle` or `SecondCycle` to build an in-memory employee set.
4. Calls `LookForRecordFirst` or `LookForRecordSecond` to filter the set by roster status, position, shift, and prior absence rules.
5. Selects the first result, loads the employee's schedule, fills the 28 date labels, and enables report generation when results exist.
6. Writes the chosen date to `Sheet1!X1`, hides Excel again, and restarts the inactivity timer.

The form uses numeric four-digit values as working shift start times. Non-numeric values are treated as roster status codes. The code distinguishes absence-type values from rest, leave, and other non-working statuses when deciding which employees and historical remarks to include.

## Employee Review

`lvListTrainee_Click` and `LoadData` display the selected employee's ID, name, position, and 28-day roster values. `LoadColor` highlights the date associated with the selected result.

`txtSearch_Change` searches the current result list as the operator types. A match becomes the selected employee and refreshes the detail display. No match produces an ETS message and clears the search field.

`cmdRefresh_Click` uses the same all-or-nothing loader as startup. On success it replaces both roster data sets, rebuilds the selectors, attempts to restore today's date, and restarts the inactivity timer. On failure it offers Retry or Cancel without first clearing the current form data.

## RFA Report Generation

```mermaid
flowchart TD
    Start["cmdPrint_Click"] --> Stop["Stop inactivity timer"]
    Stop --> Each["Process each selected employee"]
    Each --> Which{"Employee cycle?"}
    Which -->|First| Scan1["Scan backward in first-cycle roster"]
    Which -->|Second| Scan2["Scan second cycle, then cross into first cycle if needed"]
    Scan1 --> Codes["Classify prior status codes"]
    Scan2 --> Codes
    Codes --> Remarks["Combine date/status pairs into remarks"]
    Remarks --> Results["Add employee to absence report list"]
    Results --> More{"More employees?"}
    More -->|Yes| Each
    More -->|No| Write["SaveToWorksheet"]
    Write --> Layout["Clear old rows; write ID, name, position, shift, remarks"]
    Layout --> Format["Apply borders, wrapping, and row sizing"]
    Format --> Export["cmSave_Click exports Sheet1 to PDF"]
    Export --> Done["Close workbook and quit Excel"]
```

`cmdPrint_Click` does not send output directly to a printer. It builds an RFA data set by scanning backward through each employee's roster history. For a second-cycle date, the scan can continue into the prior first-cycle roster.

`SaveToWorksheet` prepares `Sheet1` as follows:

- Writes the report date and selected shift to cells in column G.
- Clears old report rows starting at row 7.
- Writes employee ID, name, position, shift, and combined absence remarks.
- Applies borders, row height, and text wrapping.
- Calls `cmSave_Click` after the sheet is ready.

`Sheet1` is the built-in report template. Its fixed column headers are `ID#`, `NAME OF EMPLOYEE`, `ROLE`, `SHIFT`, `REMARKS`, `Staff Signature`, and the two-line `Manager/Admin` / `Signature` approval header. The macros populate and format the rows beneath those headers; no separate report-template file is loaded.

`cmSave_Click` exports `Sheet1` to the configured RFA network output folder. An `All Shift` report, or any report containing more than one distinct nonblank shift, uses the date-only name `RFA Report d-mmm-yy.pdf`. A report containing exactly one shift appends that shift to the filename. An existing target PDF is replaced. If replacement or export fails, Excel is restored and the operator receives an error instead of the workbook closing silently.

## Timer and Shutdown Behavior

`Module3` provides a five-hour-and-ten-minute inactivity timer using `Application.OnTime`.

```mermaid
stateDiagram-v2
    [*] --> TimerScheduled: SetTimer
    TimerScheduled --> TimerCancelled: StopTimer
    TimerCancelled --> TimerScheduled: user/form activity completes
    TimerScheduled --> Shutdown: scheduled time reached
    Shutdown --> [*]: close target workbook / quit Excel
```

Many form events cancel and reschedule this timer. The workbook class also resets it after worksheet calculation and selection changes. `Workbook_BeforeClose` cancels the scheduled callback.

## Component Inventory

| Component | Type | Main responsibility |
|---|---|---|
| `MainWindow.frm` | UserForm source | Importable form descriptor plus UI, roster loading, selection, filtering, report creation, and PDF export code |
| `MainWindow.frx` | UserForm binary companion | Exact embedded control layout and binary properties required when importing `MainWindow.frm` |
| `ThisWorkbook.cls` | Workbook document-class source | Startup, shutdown, calculation, and selection-change event procedures for manual copying into the destination `ThisWorkbook` object |
| `Module1.bas` | Standard module | Adds minimize/maximize styles to the UserForm window through Windows APIs |
| `Module2.bas` | Standard module | Makes the UserForm appear as a taskbar application through Windows APIs |
| `Module3.bas` | Standard module | Schedules, cancels, and performs inactivity shutdown |

The retained components contain 52 active or reachable procedures and 125 detected form controls. `Sheet1`, `Sheet2`, and `Sheet3` have no executable VBA, so their empty document modules are not exported.

## Manual VBA Import

Keep `MainWindow.frm` and `MainWindow.frx` together in the same directory. In the VBA editor, import `MainWindow.frm`; Excel reads the companion `.frx` automatically, so the `.frx` is not imported separately. Then import `Module1.bas`, `Module2.bas`, and `Module3.bas` individually.

Do not import `ThisWorkbook.cls` as an ordinary class module: that would not attach its event procedures to the workbook. Open `ThisWorkbook.cls` as source, then copy its declarations and four event procedures into the destination workbook's existing `ThisWorkbook` object. The three worksheet class modules contain no executable code, but the destination must retain the `Sheet1` report template and layout.

The exported form source contains `<REDACTED>` in place of the locally embedded roster password. Replace that placeholder only in the local destination workbook. The destination also needs Microsoft ListView/Common Controls support.

## Important Procedure Groups

| Area | Procedures |
|---|---|
| Form lifecycle | `UserForm_Initialize`, `UserForm_Activate`, `UserForm_QueryClose`, `UserForm_Terminate` |
| Workbook events | `Workbook_Open`, `Workbook_BeforeClose`, `Workbook_SheetCalculate`, `Workbook_SheetSelectionChange` |
| Cycle discovery and import | `LoadStartupDataWithRetry`, `TryLoadStartupData`, `TryResolveLatestRosterPaths`, `TryParseCycleFolderName`, `FolderIsAccessible`, `TryLoadRoster`, `PopulateRosterControls`, `InitializeFilterChoices`, `SelectDefaultDate` |
| Selector reactions | `cmbDate_Change`, `cmbPosition_Change`, `cmbShift_Change`, `CallShiftStarts` |
| Main tracking | `cmdTrackNow_Click`, `FirstCycle`, `SecondCycle`, `LookForRecordFirst`, `LookForRecordSecond` |
| Available shifts | `LookForRecordFirstForShifting`, `LookForRecordSecondForShifting` |
| Employee display | `lvListTrainee_Click`, `LoadData`, `LoadColor`, `FillAll`, `ClearAll` |
| Search and refresh | `txtSearch_Change`, `cmdRefresh_Click` |
| Reporting | `cmdPrint_Click`, `SaveToWorksheet`, `cmSave_Click` |
| Window integration | `AddToForm`, `AppTasklist` |
| Timer | `SetTimer`, `StopTimer`, `ShutDown` |

The V6.0 cleanup removed 17 unreferenced procedures and five empty event handlers, along with disabled code that had been retained as full-line comments. Explanatory section and API comments were preserved. The retained procedures are either called by the application, wired to an existing form event, or used as a timer callback.

## External Dependencies

- Microsoft Excel with macros enabled.
- Windows; the project calls `user32.dll` and uses Windows-specific window handles and styles.
- Microsoft ListView/Common Controls support used by the form's ListView controls.
- Access to the `03 TG Roster` network folder, the two latest consecutive roster workbooks, and the RFA output share.
- A `Current Roster` worksheet with the layout expected by the positional ListView logic.
- A report template on `Sheet1`, including the cells used for the report date, shift, and output filename.

## Risks and Maintenance Notes

1. **Embedded credential:** A workbook/worksheet password is hard-coded in the original VBA. It should be removed from source code and supplied through an approved secure mechanism.
2. **Network coupling:** Cycle discovery, roster loading, and PDF output depend on internal network paths. There is no offline fallback.
3. **Broad Excel shutdown:** Several routines call `Application.Quit`, which can close the entire Excel instance rather than only this workbook.
4. **Legacy error handling outside startup:** Startup and roster refresh now use contextual cleanup and Retry/Cancel recovery, but some unrelated legacy handlers still suppress errors or show only a generic message.
5. **Unqualified Excel objects:** Several `Cells`, `Rows`, `Range`, and `Sheets` references rely on whichever workbook or worksheet is active.
6. **64-bit API declarations:** The declarations are marked `PtrSafe`, but window handles are stored as `Long` rather than `LongPtr`; this can be unsafe in 64-bit Office.
7. **Time-window boundaries:** Shift filters use mixed string/numeric comparisons and strict greater-than/less-than bounds, which can exclude exact boundary times.
8. **Large monolithic form:** Most business logic is embedded in `MainWindow.frm`, making testing, reuse, and isolated maintenance difficult.
9. **Static-analysis limitation:** Control bindings, `Application.OnTime`, workbook events, and any dynamically resolved calls are not fully represented by direct-call scanning.

## Export Process

The documentation and source files were produced without executing the macros:

```mermaid
flowchart LR
    XLSM["Testing workbook .xlsm"] --> VBAProject["Embedded vbaProject.bin"]
    VBAProject --> Extract["Decompress VBA source streams"]
    VBAProject --> Designer["Extract MainWindow designer storage"]
    Extract --> Keep["Write .bas, .cls, and .frm source files"]
    Designer --> FRX["Build MainWindow.frx companion"]
    Keep --> Folder["RFA/exported_vba"]
    FRX --> Folder
    Folder --> Review["Static call and workflow analysis"]
    Review --> Docs["RFA/VBA_DOCUMENTATION.md"]
```

`ThisWorkbook.cls` contains the workbook's four event procedures for manual copying into the destination document object. The empty `Sheet1`, `Sheet2`, and `Sheet3` class modules remain omitted. The UserForm binary payload is included as `MainWindow.frx`; its 90 designer streams match the embedded form byte-for-byte. Generated `manifest.json`, `PROJECT.txt`, and `references.txt` metadata files are intentionally omitted. `Testing/Macro.xlsm` remains authoritative for the report worksheet template.
