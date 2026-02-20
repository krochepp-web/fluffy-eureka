# Component Picker Intended Workflow (BOM / PO / Inventory)

This guide explains the intended end-user workflow and exactly which VBA scripts (macros) to run.

## 1) One-time setup and validation (recommended)

Run these in order before first use (or after schema changes):

1. `UI_Run_GateCheck`
2. `UI_Run_HealthCheck`
3. `UI_Run_AllChecks`
4. `UI_RefreshAutomationRegistry`

Why: these checks confirm core schema/platform readiness and refresh registry metadata before user-facing operations.

Note on Gate popups: successful Gate checks are now silent by default; failures still prompt. If `Landing` has `DEV MODE? = TRUE`, Gate PASS messages are shown.

---

## 2) Open and refresh the picker

### Default (sheet-based picker)
1. Run: `UI_Open_ComponentPicker`
2. This creates/updates `Pickers` and `TBL_PICK_RESULTS` if needed.
3. Enter filter values in:
   - `B2` search text (description / notes / PN / CompID contains)
   - `B3` revision (optional exact)
   - `B4` active only (TRUE/FALSE)
   - `B5` max results
   - `B6` CompID (optional exact match)
   - `B7` Supplier (optional exact match via dropdown)
   - `B8` Description (optional contains or wildcard match, with dropdown suggestions)
4. Run: `UI_Refresh_PickerResults` after changing filters.

### Optional UserForm launcher
- Run: `UI_Open_ComponentPicker_Form_Optional`
- If `UF_ComponentPicker` is not present, it automatically falls back to the sheet-based picker.

---

## 3) Add selected components to a target context

> In all contexts, first select one or more rows in `Pickers!TBL_PICK_RESULTS` **while the Pickers sheet is active**.
> If nothing is selected, add macros now offer: **Yes = use all displayed rows**, **No = open PN/Rev dialog (BOM flow)**, **Cancel = stop**.

### A) Add to BOM
1. Stay on `Pickers`, select one or more rows in `Pickers!TBL_PICK_RESULTS`.
2. Choose the destination BOM in `Pickers!B9` (**Target BOM**, dropdown sourced from `BOMS.TBL_BOMS[BOMTab]`).
3. Run: `UI_Add_SelectedPickerRows_To_ActiveBOM`
4. Enter default quantity when prompted.
5. Choose quantity mode:
   - **No**: apply default quantity to all selected rows
   - **Yes**: prompt for each selected row quantity

Behavior:
- Picker-driven BOM adds resolve the target BOM from `Pickers!B9`; no BOM sheet switch is required.
- If PN+Rev already exists in BOM, `QtyPer` is incremented.
- If PN+Rev is new, a row is inserted.

Fallback option (manual entry):
- Run: `UI_Add_ComponentByPNRev_To_ActiveBOM`
- Enter PN and QtyPer; Rev is optional. If left blank and multiple active revisions exist, you will be prompted to choose one.

### B) Add to PO Lines
1. Ensure `POLines!TBL_POLINES` exists and is ready.
2. Run: `UI_Add_SelectedPickerRows_To_POLines`
3. Enter default quantity and choose quantity mode.

Behavior:
- Appends rows to `TBL_POLINES`.
- Writes `CompID`, `OurPN`, `OurRev`, `Description`, `UOM`, `POQuantity` (+ audit columns when present).

### C) Add to Inventory
1. Ensure `Inv!TBL_INV` exists and is ready.
2. Run: `UI_Add_SelectedPickerRows_To_Inventory`
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

## 5) Suggested button-to-macro mapping

For easier use, assign worksheet buttons:

- **Open Picker** → `UI_Open_ComponentPicker`
- **Refresh Picker Results** → `UI_Refresh_PickerResults`
- **Add to BOM (from selected picker rows)** → `UI_Add_SelectedPickerRows_To_ActiveBOM`
- **Add to BOM (manual PN/Rev fallback)** → `UI_Add_ComponentByPNRev_To_ActiveBOM`
- **Add to PO Lines** → `UI_Add_SelectedPickerRows_To_POLines`
- **Add to Inventory** → `UI_Add_SelectedPickerRows_To_Inventory`

Optional:
- **Open Picker Form (if available)** → `UI_Open_ComponentPicker_Form_Optional`

---

## 6) Developer sanity checks after updates

Run these after changing picker logic:

1. `UI_Run_GateCheck`
2. `UI_Run_HealthCheck`
3. `UI_Run_AllChecks`
4. `UI_Run_Comps_Tests`
5. `UI_Open_ComponentPicker`
6. `UI_Refresh_PickerResults`
7. One end-to-end add test for each context:
   - `UI_Add_SelectedPickerRows_To_ActiveBOM`
   - `UI_Add_SelectedPickerRows_To_POLines`
   - `UI_Add_SelectedPickerRows_To_Inventory`

