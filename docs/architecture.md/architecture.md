# Architecture overview

This document describes the documented workflow at a conceptual level. The public repository does not contain the workbook or VBA implementation, so the flow should be validated against an approved source package before operational use.

## Documented workflow

```text
Recovery workbook
        |
        v
Excel VBA controller
        |
        +--> SAP GUI Scripting: COR3 component comparison
        |
        +--> Excel validation tables and run log
        |
        +--> SAP GUI Scripting: COR6 goods-movement processing
        |
        +--> SAP GUI Scripting: MB51 extraction
                         |
                         v
                  Excel output tables
```

## Main responsibilities

- Read batch, process-order, material, storage-location, unit, quantity, and batch values from a source recovery workbook.
- Normalize and validate the extracted values before SAP interaction.
- Compare material information with SAP COR3 data.
- Use reviewed material details to process documented COR6 phases.
- Record phase-level outcomes and error details in Excel.
- Optionally retrieve MB51 results and map activity phases back to the Excel output.

## Safety boundaries

- The public repository contains no live SAP connection code or operational workbook.
- COR3, COR6, and MB51 require an authenticated SAP GUI session and appropriate business permissions.
- COR6 Save/Post actions can change SAP production-order data and must remain subject to local operational controls.
- Synthetic data and redacted screenshots are required for public demonstrations.
- No source, configuration, or sample should contain credentials, internal paths, supplier information, production quantities, or employer-confidential rules.

## Current public status

This file is documentation for portfolio communication. It is not an executable architecture specification, and it does not replace the approved workbook's configuration, test evidence, or operating procedures.
