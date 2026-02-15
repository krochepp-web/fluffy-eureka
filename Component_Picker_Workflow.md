# Component Picker Intended Workflow (BOM / PO / Inventory)

This guide explains the intended end-user workflow and exactly which VBA scripts (macros) to run.

## 1) One-time setup and validation (recommended)

Run these in order before first use (or after schema changes):

1. `UI_Run_GateCheck`
2. `UI_Run_HealthCheck`
3. `UI_Run_AllChecks`
4. `UI_RefreshAutomationRegistry`

Why: these checks confirm core schema/platform readiness and refresh registry metadata before user-facing operations.

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
   - `B8` Description (optional exact match via dropdown)
4. Run: `UI_Refresh_PickerResults` after changing filters.

### Optional UserForm launcher
- Run: `UI_Open_ComponentPicker_Form_Optional`
- If `UF_ComponentPicker` is not present, it automatically falls back to the sheet-based picker.

---

## 3) Add selected components to a target context

> In all contexts, first select one or more rows in `Pickers!TBL_PICK_RESULTS`.

### A) Add to BOM
1. Navigate to the destination BOM sheet and ensure the BOM table is the first ListObject on that sheet.
2. Run: `UI_Add_SelectedPickerRows_To_ActiveBOM`
3. Enter default quantity when prompted.
4. Choose quantity mode:
   - **No**: apply default quantity to all selected rows
   - **Yes**: prompt for each selected row quantity

Behavior:
- If PN+Rev already exists in BOM, `QtyPer` is incremented.
- If PN+Rev is new, a row is inserted.

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
- Exact matching filters for `CompID`, `Supplier`, and `Description` when provided
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
- **Add to BOM** → `UI_Add_SelectedPickerRows_To_ActiveBOM`
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

