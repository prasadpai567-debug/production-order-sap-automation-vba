# Production Order Automation for SAP (Excel VBA)

An Excel VBA automation application that prepares production-order data, compares Excel material requirements with SAP COR3 component data, supports COR6 goods-movement entry, and records processing results in structured Excel tables.

The workbook is designed for procurement, production, and materials operations teams that need to transfer batch and material information between recovery workbooks, Excel, and SAP GUI with less manual copying.

> **Publishing note:** The original workbook contains operational data, internal file paths, company metadata, and SAP-specific implementation details. Publish only a sanitized copy. See [Security and data sanitization](#security-and-data-sanitization).

## Project overview

The application acts as an Excel-based controller and audit surface for a production-order workflow:

1. Select a completed recovery workbook.
2. Extract batch, process-order, material, storage-location, unit, quantity, and batch information.
3. Build a batch summary and material-detail table in the automation workbook.
4. Read the corresponding SAP COR3 component list through SAP GUI Scripting.
5. Match Excel materials with SAP materials and record activity phases and SAP withdrawal quantities.
6. Use the material-detail output as input for COR6 production and goods-movement processing.
7. Optionally run an SAP MB51 extraction and write the result to the `SAP Data` worksheet.

The application solves a practical data-transfer problem: production and procurement users do not need to repeatedly copy the same material, batch, quantity, and phase information between multiple Excel workbooks and SAP screens.

## Key features

### Excel recovery-workbook extraction

- Uses an Excel file picker to select a completed recovery workbook.
- Reads up to five batch numbers from the source `Entry` sheet.
- Supports `PO1`–`PO5` and zero-padded `PO01`–`PO05` worksheet names.
- Reads process-order fields, start/end dates, material numbers, material names, storage locations, units, consumption, and extracted batches.
- Carries forward material information when the source uses grouped or visually merged material rows.
- Creates or refreshes the `Batch Summary` and `Material Detail` worksheets.

### Batch and material validation

- Validates that the entry batch, PO batch, and process-order suffix agree.
- Normalizes selected storage-location aliases used by the source workbook.
- Converts a source `-` batch instruction to `FIFO*` for normal materials.
- Preserves an actual extracted batch for the fermentation-product rule.
- Uses clear comparison statuses such as `PASS`, `CHECK`, `MATERIAL MATCH`, `NOT FOUND IN SAP`, and `PRODUCT OUTPUT - NOT COMPARED`.

### SAP COR3 comparison

- Attaches to the first available SAP GUI connection and session through `GetObject("SAPGUI")`.
- Opens transaction COR3 for each process order.
- Reads the SAP component table using SAP GUI control IDs.
- Captures material number, quantity withdrawn, and activity phase.
- Aggregates repeated SAP material rows.
- Resolves the complete fermentation-product material number from the SAP material list using its configured suffix rule.
- Distinguishes numeric raw-material records from non-numeric product/output rows.
- Records SAP quantity and phase data in `Material Detail`.

### COR6 production and goods movement

- Discovers available COR6 activity phases from the SAP operation/phase overview.
- Sorts phases in ascending order.
- Reopens and verifies each phase before changing goods-movement data.
- Supports three material-entry paths:
  - actual extracted batch;
  - FIFO batch allocation from the SAP value-help list;
  - non-batch-managed material.
- Creates additional SAP rows when multiple actual batches are required for the same material and phase.
- Avoids overwriting a previously entered row when a different actual batch must remain separate.
- Handles selected SAP goods-movement error recovery steps after Save/Post.
- Skips phases that are unavailable, locked, completed, or cannot be opened, while continuing where the implementation allows it.

### Run logging

The `COR6 Run Log` worksheet records:

- timestamp;
- PO sheet;
- process order;
- activity phase;
- result;
- processing details or error information.

### SAP MB51 extraction

The `SAP DATA EXTRACTION` launcher button calls `SAP_Final_Solution`, an independent macro that:

- reads material numbers and activity phases from `Material Detail`;
- prompts for a posting-date range;
- opens SAP transaction MB51;
- enters the material selection and date range;
- filters the process order;
- exports an SAP workbook to a local temporary folder;
- normalizes the exported quantity column;
- copies the export into `SAP Data`;
- maps activity phases back to the extracted material list;
- marks unmatched activity phases as `NOT FOUND`;
- applies filters, autofit, and freeze-pane formatting.

## Application workflow

### 1. Launch the workbook

Open the macro-enabled workbook in desktop Microsoft Excel. The launcher sheet contains three worksheet shape buttons:

| Button | Assigned macro | Purpose |
|---|---|---|
| `RUN MACRO` | `Button_Recovery_Extraction_And_COR3` | Extract the recovery workbook and compare the material detail with SAP COR3. |
| `COR6` | `Button_COR6_Entry` | Process the prepared material detail through COR6. |
| `SAP DATA EXTRACTION` | `SAP_Final_Solution` | Run the MB51 extraction and populate `SAP Data`. |

The workbook does not use an automatic `Workbook_Open` workflow. The user starts processing through the launcher buttons.

### 2. Prepare the source recovery workbook

The current VBA expects the selected source workbook to contain:

- an `Entry` worksheet;
- batch numbers in column C, beginning at row 9, with `TOTAL` in column B marking the end of the batch list;
- up to five PO worksheets named `PO1`–`PO5` or `PO01`–`PO05`;
- in each PO worksheet:
  - `C5`: PO batch;
  - `I2`: process order;
  - `I3`: start date;
  - `I4`: end date;
  - `C:I`, rows 8–94: material and consumption data.

### 3. Run `RUN MACRO`

The macro opens the selected source workbook read-only, resets the output tables, extracts the available data, and writes:

- `Batch Summary`: one row per batch/PO mapping;
- `Material Detail`: one row per material/batch/phase input used for SAP comparison and COR6 processing.

### 4. Review the COR3 comparison

The macro connects to SAP GUI, opens COR3 for each process order, and writes the comparison result into `Material Detail`.

Review rows flagged as:

- `NOT FOUND IN SAP`;
- `PRODUCT OUTPUT - NOT COMPARED`;
- missing or unexpected activity phases;
- unexpected quantities or batches.

The current comparison status is based on material-number presence in COR3. The current implementation records SAP withdrawal quantity, but it does not use quantity variance as the match decision.

### 5. Run `COR6`

After the material-detail table has been reviewed, click `COR6`. The macro discovers open phases, processes them in ascending order, enters material quantities and batches, saves/posts each successful phase, and records results in `COR6 Run Log`.

COR6 can change production-order data in SAP. Use an approved SAP session and follow local operational controls.

### 6. Run `SAP DATA EXTRACTION` when required

Click `SAP DATA EXTRACTION`, enter the start and end posting dates, and allow the macro to run MB51. The extracted data is copied to `SAP Data` and enriched with activity phases from `Material Detail`.

## Example output

The following is synthetic example data for documentation only:

| PO Sheet | Entry Batch | Material No. | Activity Phase | Comparison Result |
|---|---|---|---|---|
| PO1 | BATCH-001 | RM-100001 | 0210 | MATERIAL MATCH |
| PO1 | BATCH-001 | RM-100002 | 0420 | NOT FOUND IN SAP |
| PO1 | BATCH-001 | OUTPUT-001 | — | PRODUCT OUTPUT - NOT COMPARED |

## Technical details

| Area | Implemented technology or design |
|---|---|
| Language | VBA inside an Excel `.xlsm` workbook |
| Excel features | Worksheets, shape buttons, Excel tables, file picker, InputBox, MsgBox, formatting, autofilter, freeze panes, workbook import/export |
| SAP integration | SAP GUI Scripting through the Excel/VBA COM object model |
| SAP transactions | COR3, COR6, and MB51 |
| COM helpers | `Scripting.Dictionary`, `GetObject("SAPGUI")`, `Workbooks.Open`, SAP `FindById` controls |
| Main code module | `Module3`: recovery extraction, COR3 comparison, COR6 phase and goods-movement automation |
| SAP export module | `Module4`: MB51 extraction and `SAP Data` population |
| Document modules | `ThisWorkbook` and `Sheet1` contain module attributes but no event procedures in the reviewed snapshot |
| Empty modules | `Module1` and `Module2` contain no executable procedure in the reviewed snapshot |
| UserForms | None detected. The interface is worksheet-button-driven and uses standard Excel dialogs. |
| Formulas/charts | No formulas or charts were present in the reviewed workbook snapshot. |

## Requirements

- Windows desktop environment.
- Microsoft Excel desktop with VBA and macro-enabled workbook support.
- SAP GUI for Windows with SAP GUI Scripting available and permitted by the SAP client/server configuration.
- An authenticated SAP GUI session with access to COR3, COR6, and MB51 as required by the business process.
- A source recovery workbook that follows the expected worksheet and cell contract.
- Permission to read the source workbook and write the automation workbook outputs.
- For the MB51 macro, permission to create/use the local temporary export folder configured in the VBA.

## Installation and setup

1. Download or clone the repository.
2. Obtain a sanitized, approved copy of the macro-enabled workbook. Do not use the original operational file as a public repository artifact.
3. Open the workbook in desktop Excel, not Excel for the web.
4. Confirm that the source recovery workbook and SAP GUI session are available.
5. If Excel displays Protected View, verify the file source and local security policy before enabling editing or content.
6. Open the source workbook through `RUN MACRO` when prompted.
7. Review `Batch Summary` and `Material Detail` before running COR6.
8. Run `COR6` only when the SAP session, process order, phase, material, batch, and quantity data have been reviewed.
9. Use `SAP DATA EXTRACTION` only when an MB51 output is required.

The current implementation contains environment-specific values, including an initial source-folder path, SAP GUI control IDs, an SAP selection-screen layout name, and a local export folder. These values should be treated as configuration that must be reviewed for each SAP environment.

## Business benefits

- Reduces repeated manual copying between recovery workbooks, Excel output tables, and SAP GUI.
- Standardizes batch-to-process-order validation.
- Gives users a structured material-level comparison before SAP production entry.
- Applies consistent handling for actual batches, FIFO batches, and non-batch-managed materials.
- Provides a phase-by-phase run log for operational review and troubleshooting.
- Reduces transcription risk by reusing extracted values and SAP GUI selections.
- Preserves an Excel-based audit surface that procurement and production users can review without reading VBA code.

No quantified time or error-rate improvement is claimed here because the workbook does not contain a before/after performance study.

## Known limitations and operational risks

- The automation depends on SAP GUI Scripting and recorded SAP control IDs. SAP screen or layout changes can break the automation.
- The code uses the first available SAP connection and first session rather than asking the user to select a session.
- SAP GUI must already be logged in and available to scripting.
- COR6 and Save/Post actions can change SAP production-order data. The workbook does not provide a general rollback mechanism.
- The comparison status is primarily a material-presence check. It is not a full quantity reconciliation.
- Several paths, field identifiers, and layout values are embedded in VBA rather than maintained in a user-editable configuration area.
- The source recovery workbook is an external dependency and is not included in this repository.
- The current workbook snapshot contains prior run data and error records. These are not test evidence or production-quality KPIs.

## Recommended repository structure

```text
production-order-sap-automation/
├── README.md
├── LICENSE
├── .gitignore
├── src/
│   └── ProductionOrderAutomation_Sanitized.xlsm
├── docs/
│   ├── application-analysis.md
│   ├── architecture.md
│   └── screenshots/
│       ├── launcher.png
│       ├── batch-summary.png
│       ├── material-detail.png
│       ├── cor6-run-log.png
│       └── sap-data.png
└── samples/
    └── sanitized-recovery-input.xlsx
```

### Files to publish

- `README.md` and supporting documentation.
- A sanitized macro-enabled workbook only if you are authorized to publish the VBA and workbook design.
- Synthetic sample input/output files, clearly labeled as synthetic.
- Redacted screenshots showing the workflow.

### Files to keep private

- The original operational `.xlsm` file.
- Recovery workbooks, SAP exports, run logs, material lists, process orders, supplier/customer data, and production quantities.
- Raw `vbaProject.bin` files or unreviewed `.bas` exports.
- Any internal network path, SharePoint URL, company metadata, SAP server information, or credentials.

## Suggested screenshots

Use synthetic or fully redacted data in every image:

1. Launcher sheet with the three buttons visible.
2. `Batch Summary` showing batch-to-process-order validation.
3. `Material Detail` showing material, phase, batch, quantity, and comparison-result columns.
4. `COR6 Run Log` showing processed, skipped, and error states with synthetic identifiers.
5. `SAP Data` showing an MB51-style export enriched with activity phase.
6. Optional architecture diagram showing `Recovery Workbook → Excel VBA → SAP COR3/COR6/MB51 → Excel Outputs`.

Before publishing screenshots, mask process orders, material numbers, material names, quantities, batch numbers, storage locations, dates, usernames, SAP status-bar text, file paths, and company branding.

## Future improvements

These are suggested improvements, not features claimed as implemented:

- Move source folders, SAP control IDs, transaction settings, and export paths into a reviewed configuration sheet.
- Add a dry-run mode that performs extraction and comparison without SAP Save/Post actions.
- Add a preflight check for Excel version, SAP GUI scripting, active connection/session, source-sheet contract, and writable export folder.
- Add a test harness with synthetic recovery workbooks and mocked SAP responses.
- Add explicit quantity-variance checks and tolerances to complement material-presence matching.
- Replace brittle screen-control automation with a supported SAP integration interface where available.
- Add structured error codes and a clearly separated run history rather than overwriting operational outputs.
- Add a UserForm only if a form-based interface materially improves the workflow; the current application does not contain UserForms.

## Security and data sanitization

Before publishing to GitHub:

- Remove populated `Batch Summary`, `Material Detail`, `COR6 Run Log`, and `SAP Data` values.
- Remove cached external-link data and break external links.
- Remove source-file paths and internal SharePoint/UNC locations.
- Remove company, creator, modifier, document, revision, and cloud metadata.
- Review all VBA strings for internal paths, layout names, SAP control IDs, usernames, credentials, and business rules that should remain private.
- Confirm that no credentials, API keys, tokens, passwords, or personal information are present.
- Replace operational screenshots with synthetic examples.
- Obtain employer or client approval before publishing VBA code or SAP automation logic.

## Resume and portfolio description

> Built an Excel VBA automation tool for production-order and materials operations that extracts batch and material data, compares Excel requirements with SAP COR3, supports COR6 goods-movement processing with actual-batch and FIFO logic, logs phase-level results, and automates MB51 data extraction. The project demonstrates Excel VBA development, SAP GUI automation, procurement operations, data validation, and workflow-focused problem solving.

## License

Choose a license only after confirming that you own the VBA code and have permission to publish the business logic. If the workbook was created for an employer or client, keep the operational source private and publish only an approved sanitized demonstration.
