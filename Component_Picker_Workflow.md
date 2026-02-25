# Component Picker Intended Workflow (BOM / PO / Inventory)

This guide explains the intended end-user workflow and exactly which VBA scripts (macros) to run.

Canonical BOM add entry points are:
- `UI_OP_AddSelectedPickerRowsToActiveBOM`
- `UI_OP_AddComponentByPNRevToActiveBOM`

Deprecated compatibility macro (do not map to buttons/automation):
- `DEV_LegacyAddComponentsToBOM` (wrapper that routes into picker module)

## 1) One-time setup and validation (recommended)

Run these in order before first use (or after schema changes):

1. `RunGateCheck`
2. `DEV_RunDiagnostics`
3. `DEV_RefreshAutomationRegistry`

Compatibility UI wrappers remain available (`UI_OP_RunGateCheck`, `DEV_RunHealthCheck`, `UI_OP_RunAllChecks`) but now route to the canonical check semantics.

Why: these checks confirm core schema/platform readiness and refresh registry metadata before user-facing operations.

Note on Gate popups: successful Gate checks are now silent by default; failures still prompt. If `Landing` has `DEV MODE? = TRUE`, Gate PASS messages are shown.

---

## 2) Open and refresh the picker

### Default (sheet-based picker)
1. Run: `UI_OP_OpenComponentPicker`
2. This creates/updates `Pickers` and `TBL_PICK_RESULTS` if needed.
3. Enter filter values in:
   - `B2` search text (description / notes / PN / CompID contains)
   - `B3` revision (optional exact)
   - `B4` active only (TRUE/FALSE)
   - `B5` max results
   - `B6` CompID (optional exact match)
   - `B7` Supplier (optional exact match via dropdown)
   - `B8` Description (optional contains or wildcard match, with dropdown suggestions)
4. Run: `UI_OP_RefreshPickerResults` after changing filters.

### Optional UserForm launcher
- Run: `UI_OP_OpenComponentPickerFormOptional`
- If `UF_ComponentPicker` is not present, it automatically falls back to the sheet-based picker.

---

## 3) Add selected components to a target context

> In all contexts, first select one or more rows in `Pickers!TBL_PICK_RESULTS` **while the Pickers sheet is active**.
> If nothing is selected, add macros now offer: **Yes = use all displayed rows**, **No = open PN/Rev dialog (BOM flow)**, **Cancel = stop**.

### A) Add to BOM
1. Navigate to the destination BOM sheet and ensure the BOM table is the first ListObject on that sheet.
2. Run: `UI_OP_AddSelectedPickerRowsToActiveBOM`
3. Enter default quantity when prompted.
4. Choose quantity mode:
   - **No**: apply default quantity to all selected rows
   - **Yes**: prompt for each selected row quantity

Behavior:
- If PN+Rev already exists in BOM, `QtyPer` is incremented.
- If PN+Rev is new, a row is inserted.

Fallback option (manual entry):
- Run: `UI_OP_AddComponentByPNRevToActiveBOM`
- Enter PN and QtyPer; Rev is optional. If left blank and multiple active revisions exist, you will be prompted to choose one.

### B) Add to PO Lines
1. Ensure `POLines!TBL_POLINES` exists and is ready.
2. Run: `UI_OP_AddSelectedPickerRowsToPOLines`
3. Enter default quantity and choose quantity mode.

Behavior:
- Appends rows to `TBL_POLINES`.
- Writes `CompID`, `OurPN`, `OurRev`, `Description`, `UOM`, `POQuantity` (+ audit columns when present).

### C) Add to Inventory
1. Ensure `Inv!TBL_INV` exists and is ready.
2. Run: `UI_OP_AddSelectedPickerRowsToInventory`
3. Enter default quantity and choose quantity mode.

Behavior:
- Appends rows to `TBL_INV`.
- Writes `CompID`, `OurPN`, `OurRev`, `ComponentDescription`, `UOM`, `ADD/SUBTRACT` (+ audit columns when present).

---

## 4) Validation and safety rules implemented

The picker pipeline enforces:

- Active component selection (`RevStatus = Active` when filter enabled / add processing)
- Exact matching filters for `CompID` and `Supplier`, plus contains/wildcard matching for `Description` when provided
- Positive quantity only
- Target table/header presence checks by context
- Uniqueness checks across **active** component mappings before writes:
  - duplicate active `OurPN + OurRev`
  - duplicate active `CompID`

If a check fails, the macro stops with a clear message.

---

## 5) Suggested button-to-macro mapping (canonical only)

For easier use, assign worksheet buttons:

- **Open Picker** → `UI_OP_OpenComponentPicker`
- **Refresh Picker Results** → `UI_OP_RefreshPickerResults`
- **Add to BOM (from selected picker rows)** → `UI_OP_AddSelectedPickerRowsToActiveBOM`
- **Add to BOM (manual PN/Rev fallback)** → `UI_OP_AddComponentByPNRevToActiveBOM`
- **Add to PO Lines** → `UI_OP_AddSelectedPickerRowsToPOLines`
- **Add to Inventory** → `UI_OP_AddSelectedPickerRowsToInventory`

Optional:
- **Open Picker Form (if available)** → `UI_OP_OpenComponentPickerFormOptional`

Avoid mapping legacy compatibility macro:
- ~~`DEV_LegacyAddComponentsToBOM`~~ (deprecated wrapper)

---

## 6) Developer sanity checks after updates

Run these after changing picker logic:

1. `RunGateCheck`
2. `DEV_RunDiagnostics`
3. `DEV_RunCompsTests`
4. `UI_OP_OpenComponentPicker`
5. `UI_OP_RefreshPickerResults`
6. One end-to-end add test for each context:
   - `UI_OP_AddSelectedPickerRowsToActiveBOM`
   - `UI_OP_AddSelectedPickerRowsToPOLines`
   - `UI_OP_AddSelectedPickerRowsToInventory`

