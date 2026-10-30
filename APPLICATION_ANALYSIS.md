# Application Analysis: Production Order Automation for SAP

## 1. Scope and review method

This report is based on a read-only review of the supplied macro-enabled workbook and its embedded VBA project. The review covered:

- workbook sheets, tables, buttons, defined names, external links, formulas, charts, and document metadata;
- the embedded VBA modules extracted from `vbaProject.bin`;
- launcher-button assignments and the data flow between worksheets and SAP GUI;
- the current workbook snapshot, treated only as operational sample data—not as a performance benchmark.

The original workbook was not modified, and no macro was executed during the review. Excel opened the file in Protected View, so the conclusions are based on workbook structure and source-code inspection rather than a live end-to-end SAP run.

## 2. Executive summary

The application is an Excel VBA controller for a production-order and materials workflow. It transfers structured information from a recovery workbook into Excel tables, reads material information from SAP through SAP GUI Scripting, compares the two sources, and can use the prepared data to enter goods movements in SAP COR6. A separate macro extracts SAP MB51 data and enriches it with activity-phase information.

The application addresses a manual data-transfer problem across three environments:

`Recovery workbook → Excel validation tables → SAP COR3/COR6/MB51 → Excel logs and outputs`

Its strongest implemented capabilities are structured extraction, material-level validation, SAP GUI data entry, FIFO/actual-batch handling, phase-level logging, and repeatable Excel output formatting. Its main operational risks are dependence on SAP screen control IDs, hardcoded environment-specific values, use of the first available SAP session, and live Save/Post actions without a general rollback mechanism.

## 3. What the application does

### Implemented scope

1. Lets the user choose a recovery workbook through an Excel file picker.
2. Reads batch, process-order, date, material, storage-location, unit, consumption, and batch data from the source workbook.
3. Rebuilds `Batch Summary` and `Material Detail` in the automation workbook.
4. Opens each process order in SAP COR3 and reads its component list.
5. Compares Excel material numbers with the SAP component list and records SAP quantity and activity-phase information.
6. Uses `Material Detail` to process available COR6 phases in ascending order.
7. Handles actual batches, FIFO batch selection, non-batch-managed materials, duplicate batch groups, and selected SAP goods-movement recovery scenarios.
8. Records COR6 outcomes in `COR6 Run Log`.
9. Runs an independent MB51 extraction flow and copies the result to `SAP Data`.

### Not implemented or not detected

- No UserForms were detected.
- No automatic `Workbook_Open` process was detected.
- No formulas or charts were detected in the reviewed workbook snapshot.
- No web API, database connection, Power Query process, Python component, or external automation service was detected.
- The COR3 comparison records SAP quantities but does not use quantity variance as the match decision.
- No general transaction rollback or undo mechanism is provided.

## 4. Workbook structure

| Worksheet | Role | Current design |
|---|---|---|
| `Sheet1` | Launcher | Worksheet shapes/buttons only; no normal data table. |
| `Batch Summary` | Batch-to-PO validation | Excel table `BatchSummaryTable`; stores sequence, entry batch, PO sheet, PO batch, process order, dates, and validation. |
| `Material Detail` | Main integration table | Excel table `MaterialDetailTable`; stores materials, names, storage locations, units, quantities, batches, phases, SAP quantities, and comparison results. |
| `COR6 Run Log` | Execution history | Timestamped process-order/phase/result/details log. |
| `SAP Data` | MB51 output | Populated by the separate MB51 extraction macro; empty in the reviewed snapshot before that flow is run. |

The workbook also contains an external-link relationship and cached external data. These links and cached values should be removed before public release.

## 5. Launcher buttons and entry points

The visible launcher has three worksheet shape buttons:

| Visible button | VBA entry point | Function |
|---|---|---|
| `RUN MACRO` | `Button_Recovery_Extraction_And_COR3` | Starts recovery-workbook extraction followed by the COR3 comparison. |
| `COR6` | `Button_COR6_Entry` | Validates that material data exists, then starts COR6 phase processing. |
| `SAP DATA EXTRACTION` | `SAP_Final_Solution` | Starts the independent MB51 extraction and `SAP Data` population flow. |

The source also contains a `Master_SAP_RunAll` procedure that chains COR3 and COR6 when the COR3 stage succeeds. The visible launcher assignment reviewed for the workbook uses the separate extraction/COR3 and COR6 entry points, so the master procedure should not be described as the normal button-driven path unless it is intentionally wired into the published copy.

## 6. VBA module inventory

### `Module3` — primary workflow module

This module contains the main application logic:

- source-workbook selection and extraction;
- batch summary and material-detail creation;
- storage-location, material-number, batch, and date normalization;
- COR3 SAP attachment, navigation, table reading, and comparison;
- COR6 phase discovery, phase validation, material entry, batch selection, and Save/Post handling;
- run-log creation and formatting;
- SAP session and error-handling helpers.

### `Module4` — independent MB51 extraction module

This module contains `SAP_Final_Solution` and related helpers. It drives MB51, exports the result, normalizes the quantity column, copies data into `SAP Data`, and maps activity phases from `Material Detail`.

### `ThisWorkbook` and `Sheet1`

The reviewed source contained module attributes but no executable event procedures in these modules. This supports the conclusion that the workbook is manually launched through worksheet buttons rather than automatically on open.

### `Module1` and `Module2`

No executable procedures were present in the reviewed source snapshot.

## 7. Detailed application workflow

### Stage A — select and read the recovery workbook

The user selects an `.xlsx`, `.xlsm`, or `.xlsb` source workbook. The VBA opens it read-only and expects an `Entry` worksheet plus a series of PO worksheets.

The current source contract is:

| Source area | Expected content |
|---|---|
| `Entry` column C, beginning at row 9 | Entry-batch values; the scan stops when `TOTAL` is found in column B. |
| PO sheet `C5` | PO batch. |
| PO sheet `I2` | Process order. |
| PO sheet `I3` / `I4` | Start and end dates. |
| PO sheet `C:I`, rows 8–94 | Material, description, storage location, unit, consumption, and batch fields. |

The implementation supports PO sheet names in both `PO1`–`PO5` and zero-padded `PO01`–`PO05` forms. It carries forward material attributes when the source uses grouped or visually merged rows.

### Stage B — create validation outputs

The application resets or creates `Batch Summary` and `Material Detail`, then formats them as Excel tables.

`Batch Summary` records the batch/PO relationship and assigns a validation result. The implemented mapping is `PASS` when the entry batch equals the PO batch and the process-order value ends with the entry batch; otherwise it records `CHECK`.

`Material Detail` records material-level input for the SAP comparison and COR6 flow. Source-specific normalization includes:

- selected storage-location aliases are mapped to a common storage location;
- a `-` batch instruction is converted to a FIFO marker for normal materials;
- a fermentation-product rule preserves an actual extracted batch when available;
- rows are retained when the source contains meaningful consumption or batch information.

After a successful comparison, the automation workbook saves itself. This is an operational side effect that should be documented and tested before any public demonstration.

### Stage C — read SAP COR3 and compare materials

The application attaches to SAP GUI through the Windows COM object `SAPGUI`, obtains the first available connection and first session, and opens COR3 for each process order in `Batch Summary`.

It reads the SAP component table, including:

- material number;
- quantity withdrawn;
- activity phase.

Repeated material rows are aggregated. A special product-material rule resolves the full material number from the SAP component list using a configured suffix. Numeric material keys are compared against SAP; nonnumeric output rows are marked as product output and are not compared.

The current result logic is presence-based:

- material found in SAP → `MATERIAL MATCH`;
- material not found → `NOT FOUND IN SAP`;
- nonnumeric product/output record → `PRODUCT OUTPUT - NOT COMPARED`.

Although SAP withdrawal quantity is written to `Material Detail`, the current match result is not a quantity-tolerance reconciliation. A declared comparison-tolerance constant is present in the source, but the reviewed comparison path does not use it to decide the result.

### Stage D — discover and process COR6 phases

The COR6 flow first prepares or refreshes `COR6 Run Log`, attaches to SAP, discovers selectable phases from the production-order overview, and sorts the phases in ascending order.

For each phase, the code:

1. opens COR6 and verifies that the expected phase is active;
2. finds the goods-movement table;
3. selects rows in `Material Detail` for the current PO sheet and activity phase;
4. groups matching material/phase/storage combinations;
5. enters quantities and batch values using the appropriate batch-management path;
6. allocates FIFO quantities from SAP batch selection when required;
7. creates additional rows when multiple actual batches are needed;
8. saves/posts the phase when the SAP response is acceptable;
9. records the result and details in `COR6 Run Log`.

The three material-entry paths are:

- actual extracted batch;
- FIFO or system-selected batch allocation;
- non-batch-managed material.

The code also recognizes unavailable, completed, locked, or non-openable phases and can log them as skipped while continuing with other phases. It checks status-bar messages and unexpected popups and contains a specific recovery sequence for one goods-movement error pattern.

### Stage E — optional MB51 extraction

The separate `SAP DATA EXTRACTION` entry point:

1. reads material numbers from `Material Detail`;
2. derives the relevant process-order value;
3. prompts for a posting-date range;
4. opens MB51 and fills the selection screen;
5. executes the report and filters the result for the process order;
6. exports the result to a local temporary folder;
7. normalizes quantity values, including negative values and comma formatting;
8. copies the export into `SAP Data`;
9. adds activity phase values using a material-to-phase dictionary;
10. formats the result with filtering, autofit, and frozen headers.

## 8. Current workbook snapshot findings

These figures describe the attached workbook at review time only. They are not test results, production KPIs, or claims about the application’s general accuracy.

- `Batch Summary`: 3 data rows; 2 `PASS` and 1 `CHECK`.
- `Material Detail`: 22 data rows; all current rows show `MATERIAL MATCH`.
- `COR6 Run Log`: 53 data rows; 14 `Processed`, 32 `No Materials Found`, and 7 `Error` results.
- The log contains multiple activity phases and prior-run operational details.
- The workbook contains populated operational values and should not be uploaded to a public repository in its current form.

The presence of error records is important: a clean README should describe the workflow and known states, not imply that every phase in the attached snapshot completed successfully.

## 9. Technical stack and dependencies

| Category | Finding |
|---|---|
| Programming language | VBA embedded in a macro-enabled Excel workbook. |
| Host application | Microsoft Excel desktop on Windows. |
| SAP integration | SAP GUI for Windows through SAP GUI Scripting and COM. |
| SAP transactions | COR3 for component comparison, COR6 for goods movement, and MB51 for material-document extraction. |
| Excel features | Worksheets, Excel tables, shapes, file picker, InputBox, MsgBox, filtering, formatting, freeze panes, workbook open/save/close. |
| VBA helpers | `Scripting.Dictionary`, standard Excel/VBA object model, SAP `FindById` and table/grid controls. |
| External dependency | A source recovery workbook following the expected sheet/cell contract. |
| Configuration model | Several source folders, SAP control identifiers, selection-layout values, and export locations are embedded in code. |
| Network/API status | No web API or database integration was found. The SAP connection is through the locally running SAP GUI client. |

The source recovery workbook and SAP GUI configuration are part of the runtime environment but are not appropriate repository files.

## 10. Business benefits

The implemented design can provide the following operational benefits:

- reduces repeated manual copying between recovery workbooks and SAP;
- creates a consistent batch-to-process-order validation step;
- exposes material-level differences before COR6 data entry;
- standardizes handling of actual batches, FIFO batches, and non-batch materials;
- processes COR6 phases in a repeatable order;
- retains phase-level results and error details for review;
- reduces transcription risk by reusing extracted values;
- gives non-developers an Excel-based output and review surface.

No numerical time savings, error-rate reduction, or return-on-investment claim should be published until supported by measured before/after evidence.

## 11. Limitations and operational risks

1. **SAP screen coupling:** SAP GUI control IDs and screen layouts can change, breaking the automation.
2. **Session selection:** the implementation uses the first SAP connection and first session it finds, which is risky on machines with multiple sessions.
3. **Live SAP mutations:** COR6 Save/Post actions can change production-order data; the workbook does not provide a general rollback.
4. **Presence rather than quantity matching:** COR3 comparison does not currently make quantity variance part of the match decision.
5. **Hardcoded configuration:** source folder, export folder, SAP control identifiers, and layout values are not fully externalized.
6. **Input contract:** source workbook layout and naming conventions are assumed rather than validated through a formal schema screen.
7. **Operational overwrite:** output sheets are reset and the workbook saves after successful comparison, so users should archive outputs deliberately.
8. **Limited testability:** the code is tightly coupled to live Excel and SAP GUI state; no mock SAP test harness was found.
9. **Error continuation:** skipped and error phases may allow later phases to continue, so users must review the complete run log rather than relying on the final message alone.

## 12. Security and publication review

### Data that should be removed or replaced

Before publishing, remove or replace:

- batch numbers, process orders, material numbers, material names, quantities, storage locations, dates, activity phases, and SAP error text;
- populated `Batch Summary`, `Material Detail`, `COR6 Run Log`, and `SAP Data` worksheets;
- cached external-link data and external-link targets;
- internal network, SharePoint, UNC, recovery, or user-profile paths;
- SAP server/session details, control IDs, layout names, and business rules that are confidential to the employer;
- company branding and document metadata;
- creator and last-modified-by values, revision identifiers, cloud/recovery metadata, and other hidden properties;
- raw `vbaProject.bin`, `__SRP_*` cache streams, or unreviewed source exports.

Use synthetic workbook data and redacted screenshots for the public demonstration.

### Credential check

No credentials, passwords, API keys, or access tokens were identified in the reviewed VBA strings. This is not a security guarantee: perform an independent scan of the sanitized copy, workbook relationships, custom document properties, and repository history before publication.

### Publishing recommendation

Publish documentation and synthetic examples first. Publish the macro-enabled workbook or VBA source only if the author has permission from the employer/client and has completed a code, metadata, external-link, and data-privacy review.

## 13. Recommended repository structure

```text
production-order-sap-automation/
├── README.md
├── LICENSE
├── .gitignore
├── src/
│   └── ProductionOrderAutomation_Sanitized.xlsm
├── docs/
│   ├── application-analysis.md
│   └── screenshots/
│       ├── launcher.png
│       ├── batch-summary.png
│       ├── material-detail.png
│       ├── cor6-run-log.png
│       └── sap-data.png
└── samples/
    └── sanitized-recovery-input.xlsx
```

### Recommended public files

- `README.md`;
- this analysis report, renamed to `docs/application-analysis.md`;
- synthetic source and output samples;
- redacted screenshots;
- an approved sanitized workbook or reviewed VBA source;
- a license only after publication rights are confirmed.

### Keep private

- the original operational workbook;
- real recovery workbooks and SAP exports;
- actual logs, process-order data, material master data, supplier/customer data, and production quantities;
- raw macro project binaries and unreviewed source dumps;
- internal configuration values and credentials, if any are found during the final scan.

## 14. Recommended screenshots

Use synthetic or fully masked values in every image:

1. Launcher sheet showing the three available actions.
2. `Batch Summary` showing a successful and a review-required validation state.
3. `Material Detail` showing comparison results and activity-phase fields.
4. `COR6 Run Log` showing processed, skipped, and error records.
5. `SAP Data` showing MB51-style output enriched with activity phase.
6. Optional architecture diagram showing the recovery workbook, Excel VBA, SAP transactions, and Excel outputs.

Mask process orders, batch numbers, material numbers/names, quantities, dates, storage locations, paths, usernames, status-bar text, company names, and SAP server information.

## 15. Future improvements

The following are recommendations, not existing features:

- move paths, SAP identifiers, and export settings to a reviewed configuration sheet;
- add a preflight screen for Excel version, SAP scripting availability, active session, source layout, and writable folders;
- add a dry-run mode that never performs SAP Save/Post;
- implement quantity-variance checks with an explicit tolerance and separate status;
- add a mock SAP/test harness using synthetic recovery workbooks;
- add explicit error codes and a durable, append-only run history;
- allow the user to choose a specific SAP connection/session;
- replace brittle screen automation with a supported SAP integration interface where available;
- add controlled versioning and release notes for changes to the source workbook contract.

## 16. Portfolio description

> Built an Excel VBA automation tool for production-order and materials operations that extracts batch and material data, compares Excel requirements with SAP COR3, supports COR6 goods-movement processing with actual-batch and FIFO logic, logs phase-level results, and automates MB51 data extraction. The project demonstrates Excel VBA development, SAP GUI automation, data validation, and workflow-focused problem solving.

This description intentionally avoids claiming quantified savings or full quantity reconciliation because those capabilities were not evidenced by the reviewed application.

