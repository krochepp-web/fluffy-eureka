<u>010 - Documentation Strategy</u>

010 - Documentation Strategy

This document.

020 - Project Charter

high-level definition of What this workbook will do and generally How it will be used

030 - Custom GPT Instructions

Instructions for a custom GPT geared towards achieving the goals outlined in this charter.

040 - Vibe Coding Guidelines

best practices guide to help specify the coding format and strategy, both for the user and the ai coding assistant

050 - Workbook Layout & Functional Strategy

specifics on how every column is named, the case to use in all circumstances, core columns identified

055 – Workbook Schema Definition v3.2.0

100 - Tracker Workbook

the end product

The goal of the documentation strategy is to indicate the full documentation package that shall (ultimately) specify the complete workbook.

Global Note: This workbook will be developed by one developer with the assistance of a custom GPT as identified in 030.

1.  The Excel Workbook Component Tracker will be used to keep track of the following information:

    1.  Suppliers – a list of suppliers who can provide purchased components

    2.  Components – a list of components we may need to procure to build devices

    3.  Bill of Materials (BOM) – a list of components, and their quantities, necessary to build an assembly

    4.  Work Orders (aka Builds) – a list of commitments to provide assembled devices to the customer

    5.  Demand – a list of components needed to satisfy builds; net demand is defined as quantity needed minus quantity on-hand

    6.  Purchase Orders – a list of orders placed with suppliers to obtain the components in demand

    7.  Inventory – updated when components arrive and are consumed

    8.  Log

2.  Excel 2016+, deployed on O365 SharePoint is the compatibility target.

3.  Multiple team members will review and edit the workbook without causing data-integrity or usability issues.

4.  Each list above will reside in a table, one table per workbook tab.

5.  The golden path below describes the standard workflow for maintaining the workbook.

6.  Simplicity of functionality, ease of use, error-prevention, avoiding redundant data-entry are the primary goals.

7.  The workbook will have formulas for real-time demand; VBA orchestration for downstream workflows.

<!-- -->

1.  Buyer (user) converts demand into POs via automated button-driven workflow.

2.  Buyer or designee (user) updates the PO Tracker manually to reflect receipts, update the inventory and component demand.

3.  Changes made by users will be logged for error checking and change tracking / control via a LOG sheet containing TBL_LOG including timestamps & user IDs.

4.  Standardized inputs and validations will minimize user error.

The “Golden Path”:

1.  **Designer (user) -** defines suppliers, components, and a top-level BOM.

2.  **Project Manager (user) -** enters build commitments, which automatically generate component demand.

3.  **Buyer (user)** - reviews demand, converts required items into purchase orders through an automated workflow, and updates receipts as components arrive.

> Role mapping: Designer / PM / Buyer = User; Admin = Developer in TBL_USERS.
>
> **Outcome:**

- The team shares a single, error-resistant workbook showing assemblies required, components needed, and procurement status

- TBL_USERS will govern user-specific permissions \[admin & user & viewer to begin – the admin is the workbook developer\]

Global Note: This workbook will be developed by one developer with the assistance of a custom GPT as identified in 030.

030 - Custom GPT Instructions: Excel Automation

Purpose:

Provide robust, copy paste ready VBA and workbook guidance that survives long chats and prevents “context fade.”

Align deliverables to a single, enforceable style; keep the workbook auditable, testable, and safe for Excel 2016+ (32/64 bit).

Top Level Objectives:

1\) Produce full, runnable code (never partials) that minimizes post paste edits.

2\) Favor readability, predictability, and maintainability over clever tricks.

3\) Keep behavior deterministic across Excel 2016-O365; avoid platform specific APIs by default.

4\) Treat the Workbook Schema (tabs, tables, headers, named ranges) as the ground truth.

Always Do (Default Behavior):

1\) Deliver complete code blocks: either a full module (with Option Explicit) or a full Sub/Function.

2\) Prepend a top comment header to every code block:

\- Purpose

\- Inputs (tabs/tables/headers)

\- Outputs / Side effects

\- Preconditions / Postconditions

\- Errors & Guards

\- Version (semver), Author, Date

3\) No Select/Activate. Use fully qualified references (ws, lo, rng) and structured references by header name, not column index.

4\) Declare everything: Option Explicit; variables explicitly typed; Long not Integer; Const for literals; no magic numbers.

5\) Error handling pattern (no blanket Resume Next):

\- On Error GoTo EH

\- CleanExit label

\- EH: LogEvent + friendly MsgBox + Resume CleanExit

6\) Cite source logic when non-obvious (e.g., “MSDN: Range.Find options”); do not add external links.

7\) Compatibility: assume 32/64 bit; avoid Windows API unless explicitly requested; prefer pure VBA/worksheet functions.

Naming & Style Conventions:

\- Procedures/Public Functions: Verb_Noun (underscore), verbs first (Generate_POLine, Update_DemandTotals).

\- Module names (.bas): M\_ (M_Demand_Recalc, M_PO_Generate, M_Inventory_Receive).

\- Core modules: M_Core_Utils, M_Core_Constants, M_Core_Types, M_Core_Logging, M_Core_Schema.

\- Variables: camelCase (buildQty, poId) with helpful object prefixes (ws, wb, lo, rng, dic).

\- Constants: SCREAMING_SNAKE_CASE (DEFAULT_UOM, ERR_NO_SUPPLIER).

\- Sheets (tabs): TitleCase (Components, WOS, Demand, PO_List, Inv, Log).

\- Tables: TBL\_ (TBL_SUPPLIERS, TBL_COMPS, TBL_SCHEMA, TBL_AUTOMATION, TBL_HELPERS, TBL_BOM\_\[TA PN\], TBL_BOMS, TBL_WOS, TBL_DEMAND, TBL_PO_LINES, TBL_PO_HEADERS, TBL_INV, TBL_LOG, TBL_USERS, TBL\_-other-).

\- Named ranges: NR\_ (NR_UnitsOfMeasure, NR_DefaultLeadDays).

\- Buttons/Shapes: BTN\_ (BTN_Generate_PO, BTN_Recalc_Demand).

Workbook Schema (Required Contract):

Examples of required core tables include TBL_SUPPLIERS, TBL_COMPS, TBL_SCHEMA, TBL_AUTOMATION, TBL_HELPERS, TBL_BOM\_\[TA PN\], TBL_BOMS, TBL_WOS, TBL_DEMAND, TBL_PO_LINES, TBL_PO_HEADERS, TBL_INV, TBL_LOG, TBL_USERS. Their detailed structure is defined in the Workbook Strategy (050).

\- Provide (or request) a machine-readable schema (CSV or JSON) that lists:

\- Tabs; Tables; Column headers (exact text); Named ranges; Data validation lists; Primary keys.

\- GPT must validate the schema at the start of each module via M_Core_Schema.ValidateSchema().

\- If a required tab/table/header is missing or renamed, fail fast with a clear message and log entry.

Data Access & Joins:

\- Never rely on ActiveCell; resolve objects explicitly (ws, lo).

\- Use ListObjects and headers for all column access (lo.ListColumns("PartNumber")).

\- Build dictionaries keyed by business keys (e.g., PartNumber) for joins (BOM → Components) to avoid nested loops.

\- Avoid volatile worksheet functions unless explicitly needed; keep formulas simple and documented next to the cell/range.

When to Use What (Decision Policy):

\- Formulas: transparent, cell local math that must live update (e.g., NetDemand = Needed - OnHand).

\- Data Validation: constrained inputs from named ranges or table columns. No hard coded lists.

\- VBA: orchestration across sheets, bulk updates, PO generation, reconciliation, logging, and guarded operations.

\- Default approach: formulas for visibility; VBA for orchestration/IO; validation for controlled inputs.

Performance & Stability:

\- Standard toggles at procedure start/end (with finally block safety):

\- Application.ScreenUpdating, Application.EnableEvents

\- Application.Calculation set to xlCalculationManual within scope; restore original on exit

\- Avoid UsedRange pitfalls; qualify loops to DataBodyRange where possible.

\- Prefer arrays and dictionary lookups over cell by cell loops when processing thousands of rows.

\- Provide a DryRun As Boolean parameter for destructive actions; default False, but include a Test harness that uses True.

Error Handling & Logging:

\- Central LogEvent(procName, errNum, errDesc, optional details) appending to TBL_LOG on Log sheet with timestamp + user (Environ("Username")).

\- Validate inputs early: missing keys, negative/zero quantities, duplicates, blanks in required fields, mismatched UoM, etc.

\- Guard external I/O (file, SharePoint) with explicit checks and helpful messages.

Testing & “Idiot Proofing”:

\- Provide a Test\_() harness for each public action with reversible or mock operations.

\- Include small mock data seeds in comments or a dedicated TBL_MOCK\_\* table where useful.

\- Add ConfirmProceed() for operations that delete or overwrite data.

\- Where ambiguity exists, present 2 options with trade offs; do not guess silently.

UI & Buttons:

\- Each public action gets an optional button with a clear label. Buttons call only the public entry point.

\- Button names: BTN\_; assign macros to the public Sub only (no helpers).

\- Tooltips/notes near buttons: 1–2 bullets (what it does, what it touches).

Versioning & Change Control:

\- SemVer in headers (e.g., v1.2.0) and a tab level Change Log (TBL_CHANGELOG) with Date, Module, Version, Summary, Author.

\- Embed a constant APP_VERSION in M_Core_Constants; surface it in a small About dialog (optional).

Collaboration & Environments:

\- Preferred file format: .xlsm for macro workbooks; consider .xlsb for size/performance once stable.

\- For SharePoint/OneDrive: caution users about simultaneous editing; advise “Check Out” for macro edits.

\- Macro Security: recommend signing once stable; store trusted location guidance in comments.

Review & Tooling:

\- Encourage exporting modules (bas/.cls/.frm) to text for diffing and code review.

\- If the user has Rubberduck VBA, note where inspections would flag things (but do not require it).

Response Format (Back to User):

1\) One paragraph summary of what the code does and where it plugs in (tabs/tables/columns).

2\) Full, paste ready module/procedure (no ellipses).

3\) Assumptions & required names (tabs, tables, headers).

4\) How to run it (button assign or macro call).

5\) What to test (1–3 checks).

6\) If ambiguous, present 2 options + trade offs and ask which to implement.

Code Contract Matrix (Fields) contained in TBL_AUTOMATION on the Auto Tab:

\- Feature

\- Trigger (User/Auto)

\- Inputs (Tables/Cols)

\- Outputs / Side effects

\- Public Entry Point

\- Helpers

\- Errors / Guards

\- Tests

\- Notes / Version

Appendices:

\- Appendix A: Minimal Core Modules (M_Core_Constants, M_Core_Logging, M_Core_Schema, M_Core_Utils)

\- Appendix B: Excel 2016 Compatibility Rules

\- Appendix C: AI Collaboration Prompts

Global Note: This workbook will be developed by one developer with the assistance of a custom GPT as identified in 030.

## 040 – Vibe Coding Best Practices

## Purpose

Enable efficient, creative, and sustainable collaboration with AI while keeping the codebase understandable, testable, and maintainable. This document governs the style and workflow of human + AI development inside the Tracker Workbook.

## Core Tenets (The Non-Negotiables)

- Clarity Over Cleverness: Readability wins. Expand logic into named steps. No “tricky one-liners” that future contributors cannot understand.

- Small Iterations, Fast Feedback: Work in short cycles: state intent → generate draft → test quickly → refine. Avoid giant multi-feature drops.

- Schema Is Law: Workbook schema (tabs, tables, headers, keys) is the single source of truth. Code that doesn’t validate against schema fails by design.

- Error-Friendly Development: Assume mistakes will happen (AI or human). Always include error handling, dry-run toggles, and logging into TBL_LOG.

- Consistency Is the Currency: Follow naming, indentation, error-handler, and module header standards. Consistency buys maintainability, trust, and long-term usability.

## Common Hiccups & How to Avoid Them

- Overly Clever Code  
  Symptom: One-liners or opaque logic using With, chained conditions, or hidden defaults.  
  Fix: Expand into explicit variables and header-based references. Comment intent, not mechanics.

- Context Fade in AI Sessions  
  Symptom: GPT forgets schema, produces code with missing/incorrect headers.  
  Fix: Restate schema, assumptions, and goals at the start of every session. Use prompts that anchor back to 030/050 rules.

- Schema Drift  
  Symptom: Silent breakages from column renames (“Qty” vs “Quantity”).  
  Fix: Run M_Core_Schema.ValidateSchema before any module runs. Update schema doc first, then code.

- Testing Discipline Slips  
  Symptom: Macro works on one dataset but fails in production.  
  Fix: Always run against TBL_TEST_CASES with golden datasets. Expand test suite as new features are added.

- Unreviewed AI Code  
  Symptom: Draft code pasted into production without human oversight.  
  Fix: AI drafts, human locks. Every procedure must carry author, date, and version. Use TBL_AUTOMATION to track ownership.

- Version Confusion  
  Symptom: Users don’t know which iteration of a macro is in use.  
  Fix: Increment SemVer in module headers. Log changes in ABOUT sheet and optionally TBL_CHANGELOG.

## Workflow (Golden Path)

1.  State Intent – Define feature, function, or fix, including schema references.

2.  Generate Draft – AI scaffolds the code.

3.  Review/Challenge – Human checks for schema alignment, readability, and correctness.

4.  Refine/Test – Run small tests against TBL_TEST_CASES; correct issues fast.

5.  Lock Version – Export module, bump SemVer, update TBL_AUTOMATION.

## Outcome

A natural, collaborative coding flow that balances speed and creativity with quality. The workbook remains:  
- Consistent (same naming, error handling, versioning).  
- Maintainable (any contributor can pick up where another left off).  
- Trustworthy (schema-enforced, test-driven, and logged).

Global Note: This workbook will be developed by one developer with the assistance of a custom GPT as identified in 030.

# Contents

[Workbook Creation Strategy V3.4.1 [2](#workbook-creation-strategy-v3.4.13.4.2)](#workbook-creation-strategy-v3.4.13.4.2)

[*Executive Summary* [2](#executive-summary)](#executive-summary)

[Purpose & Scope [2](#purpose-scope)](#purpose-scope)

[Data Model (Relational Backbone) Rules: [3](#data-model-relational-backbone-rules)](#data-model-relational-backbone-rules)

[Guardrails & Validation [4](#guardrails-validation)](#guardrails-validation)

[Performance & Stability [4](#performance-stability)](#performance-stability)

[Roles & Permissions (Operational) [5](#roles-permissions-operational)](#roles-permissions-operational)

[Automation Registry [5](#automation-registry)](#automation-registry)

[Golden Paths & Edge Cases [7](#golden-paths-edge-cases)](#golden-paths-edge-cases)

[Testing Strategy (Starter Kit) [9](#testing-strategy-starter-kit)](#testing-strategy-starter-kit)

[Iteration & Deliverables [9](#iteration-deliverables)](#iteration-deliverables)

[Versioning & Change Control [9](#versioning-change-control)](#versioning-change-control)

[Workbook Schema [9](#workbook-schema)](#workbook-schema)

[SemVer in module headers and a small ABOUT cell block. [9](#semver-in-module-headers-and-a-small-about-cell-block.)](#semver-in-module-headers-and-a-small-about-cell-block.)

*Release Lock & Archive Procedure (v3.4.1)  
For each release, increment the workbook SemVer (Workbook + Doc 050) and freeze the corresponding Schema version. Run UI_Run_HealthCheck and confirm PASS (Schema_Check and Data_Check show zero issues). Run UI_RefreshAutomationRegistry to ensure Auto.TBL_AUTO reflects the current public procedure inventory and flags any removed procedures as STALE. Save the workbook as CUSTOM_TRACKER\_\<version\>.xlsm. Export all VBA modules using Export_All_VBAModules_SaveAsFolder into a versioned folder, and store a copy of Doc 050 in the same folder. Finally, zip the folder (workbook + .bas exports + documentation) to create a rollback snapshot for auditability and recovery.*

**  **

# <u>Workbook Creation Strategy V3.4.1🡪<span class="mark">3.4.2</span></u>

## *Executive Summary*

This document is the blueprint for the Tracker Workbook. It fixes the schema (tabs, tables, columns, keys), coding rules, guardrails, test approach, and iteration cadence. It aligns with the 010 Documentation Strategy, 020 Charter, 030 Custom GPT Instructions, and 040 Vibe-Coding Best Practices.

## Purpose & Scope

Purpose

Lock a **schema-first**, **auditable** workbook with deterministic VBA orchestration and table-driven tests.

Scope

The workbook includes core tables (e.g., TBL_SUPPLIERS, TBL_COMPS, TBL_WOS, TBL_POS, TBL_INV) for managing supply chain data.

See <u>Section [0](#workbook-schema)</u> for a reference to the current workbook schema definition document which is an important supplement to this strategy document.

Design Principles

Suggested best practices for documentation:

**Document** BEFORE building. Every module must have a header block in the documentation before its first line of code is written.

**Document** contracts, not internals. One example per module, and make it realistic — e.g., “Generate_WOComps expands a build of 10 assemblies with a BOM of 4 components.”

**Include** “why” for decisions. Dict lookups avoid nested loops and ensure deterministic performance even for 10k-row BOMs.

Keep everything versioned. **Always** tag:

- Workbook version

- Schema version

- Module version

- Change date

This supports later audit and backwards compatibility.

**Maintain** several “one-pagers”

- One-page Workflow Overview

- One-page Table Responsibility Grid

- One-page Modules & Ownership

- One-page Core Concepts (keys, fks, schema)

BOM structure is intentionally mutable; Work Orders always reference the current BOM state, and demand recalculates accordingly. This behavior is deterministic and auditable.

Data Structure

- **Tables** only (ListObjects); no free-range cells.

- Each table includes **audit columns** where indicated: Created/Updated At/By.

- Columns must match TBL_SCHEMA exactly (case sensitive header text). No abbreviations or synonyms permitted once Schema is frozen.

Naming conventions & syntax standardization:

- Sheets (Tab names), examples: Suppliers, Comps, Users, WOS, POS, Inv

- Tables all caps, examples: TBL_SUPPLIERS

- Columns, examples: PascalCase (LeadTimeDays)

- Modules, prefix by layer, examples: M_Core\_\*, M_Biz\_\*,

- <span class="mark">Convention**:** infra = M_Core\_\*; features = M_Biz\_\*,</span>

- Procedures Verb_Noun, examples: Create_NewComp()

- <span class="mark">Buttons, examples: BTN\_ \*, BTN_New_PO</span>

- Named Ranges, examples: NR_UnitsOfMeasure, NR_DfaultLeadDays

Workbook Tab classes

**Production tabs (user-facing):** contain user workflows, controlled data-entry fields, and reporting.

**System tabs (hidden/protected):** SCHEMA, AUTO, Log, Schema_Check, Data_Check (and similar). These support validation, logging, and automation.

**Developer tabs (hidden/protected, non-shipping optional):** e.g., Dev_ModuleInventory, internal inventories, staging artifacts.

**Visibility policy**

Non-production tabs may be hidden and protected in released versions.

End users are not permitted to modify schema, macros, validation logic, or records outside designated data-entry fields.

Workbook protection model

Workbook is distributed as a locked system. Users interact via:

- pre-defined data-entry fields

- buttons that execute approved macros

- reports/views

- All other ranges are protected.

Macros are not user-modified; development occurs only by the designated developer.

Code Behavior

Deterministic code: No Select/Activate, explicit objects, header-based access, centralized logging/validation.

Versioning: SemVer in module headers and **ABOUT** sheet

Testing: table-driven test harness with golden dataset.

Users run UI macros (only UI\_\*); **UI macros are parameterless** and worker procs should be Private where possible.

All operational macros are protected by Gate

Registry + schema + health checks are the “platform”

Operational Gates

All operational macros must call the workbook Gate before executing write actions.

The Gate blocks execution unless:

> Schema validation passes (tables/columns match SCHEMA!TBL_SCHEMA)
>
> Data Integrity validation passes (required fields, uniqueness, and foreign-key existence as defined in SCHEMA!TBL_SCHEMA)

Gate results are written to Schema_Check and Data_Check outputs and are logged to Log.

## Data Model (Relational Backbone) Rules:

- Every entity is a ListObject with an enforced key.

- FK columns refer to keys of parent tables.

- No duplicates; no negative quantities unless explicitly allowed.

- All tables to include CreatedAt, UpdatedAt, CreatedBy, UpdatedBy for auditing purposes.

See M_Auto\_<span class="mark">Seed</span> for current automation registry entries <span class="mark">(outdated for v3.4.2?).</span>

Script details and metadata are maintained in TBL_AUTO, populated by *M_Core_Automation*

Workbook Schema: See below; schema defined in document 055.

## Guardrails & Validation

On-sheet: Data Validation for enums, ranges, required fields from TBL_HELPERS/named ranges.

Pre-save: Run *Schema_Validate_All* checks (unique keys, FK existence, no negatives, required fields).

Missing tables, Extra tables, Missing columns, Extra columns, Datatype token mismatches, FK targets

Output location (Schema_Check tab)

Blocking logic: Demand or PO creation fails if IsBuildable = N <span class="mark">\[only if we decide to allow buildable=no records?\]</span>

Guard macros:

- *Schema_Validate_All* (): verify required tables/columns.

- RevCheck(): enforce workbook version compatibility.

- ConfirmProceed(): prompt for destructive actions.

Gate-checker:

- Gate_Ready(showUserMessage As Boolean) As Boolean

Depends on:

- Schema_Validate_All output sheet Schema_Check

- Validate_DataIntegrity_All output sheet Data_Check

Used by:

- Health check and eventually all “operational” macros

Error Handling & Logging

- Logging: All errors/warnings/events → TBL_LOG with timestamp & user.

Standard pattern:

> On Error GoTo EH  
> '...  
> CleanExit:  
> Exit Sub  
> EH:  
> M_Core_Logging.LogEvent "ProcName", Err.Number, Err.Description  
> Resume CleanExit

No custom Err_Log; use M_Core_Logging.LogEvent only.

Pre-action checks FK existence, unique keys, non-negative quantities, IsBuildable gating.

TYPES of ERRORS: LogEvent / LogInfo / LogWarn / LogError

TBL_LOG structure

ActivityId concept

Rules (non-recursive, safe failure, version capture)

## Performance & Stability

- Batch read/write via ListObject.DataBodyRange.Value.

- Use arrays/dictionaries for joins; avoid volatile functions on data tables.

- Scoped toggles for ScreenUpdating, Calculation, EnableEvents.

- No volatile worksheet functions on data tables.

Intended utility scripts for robust automation (this is incomplete – consider expanding in 3.4.1+):

M_Core_Utils – Planned Responsibilities

- ConfirmProceed

- SafeGetListObject

- Safe column operations

- Dictionary builders (PK, PN/Rev)

- Message utilities

- ActivityId function

- Array helpers

M_Core_Toggles – Planned Responsibilities

- Turning ScreenUpdating off/on

- Disabling/Restoring Events

- Manual/Automatic calc switching

> **Core (M_Core\_\*)**: framework + enforcement (Schema, DataIntegrity, Gate, Lockdown, Logging, Toggles, Utils), contains infrastructure and *no* button entry points except perhaps UI_Run_HealthCheck if you want it “core”
>
> **UI (M_UI\_\*)**: only button-safe entry points (no business logic); **contains only UI\_ procedures**
>
> **Feature (M_Data\_\* / M_Feature\_\*)**: workflows (Supplier entry, Components, PO generation, etc.)
>
> **Dev (Dev\_\*)**: exports, reports, workbook inventory generation, contains only Dev\_ procedures
>
> **Test (Test\_\* / M_Core_Tests)**: harness procedures only
>
> **Legacy (Legacy\_\*)**: temporary quarantine before deletion

## Roles & Permissions (Operational)

Roles stored in TBL_USERS: Developer, User (Buyers, PM’s, QA), Viewer (Any)

- **Developer** owns schema, automation, and testing.

- **Users** own correctness of data-entry tabs.

- **Viewer** is read-only.

- **Admin** is only one who can change supplier to inactive

Gating: Buttons/macros check role before proceeding.

Protection: Data tabs unlock only data body cells; Config/Schema tabs fully locked.

## Automation Registry 

The Automation Registry (TBL_AUTO) is the *code contract matrix*.

Track every public procedure with metadata:

> '\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*
>
> ' Module:      M\_\<ModuleName\>
>
> ' Procedure:   \<ProcedureName\>
>
> ' Purpose:
>
> '   \<1–3 sentences describing what the procedure does and why it exists.\>
>
> ' Inputs (Tabs/Tables/Columns):
>
> '   - \<ListObject or sheet relied upon\>
>
> '   - \<Columns accessed, by name, not index\>
>
> ' Outputs / Side Effects:
>
> '   - \<Tables updated or created\>
>
> '   - \<Logs written to TBL_LOG\>
>
> '   - \<Events triggered or follow-on calculations\>
>
> ' Preconditions:
>
> '   - Schema_Validate_All must pass with no critical errors.
>
> '   - Required tables/columns exist as defined in TBL_SCHEMA.
>
> '   - User has appropriate role (Developer/User/Viewer).
>
> ' Postconditions:
>
> '   - \<What is guaranteed true after execution\>
>
> ' Errors & Guards:
>
> '   - Logs errors using M_Core_Logging.LogEvent.
>
> '   - Fails fast on missing keys, invalid FK relationships, negative quantities,
>
> '     or broken schema (MissingColumn/MissingTable).
>
> ' Version:     vMAJOR.MINOR.PATCH
>
> ' Author:      \<Name\>
>
> ' Date:        \<YYYY-MM-DD\>
>
> ' @spec
>
> '   Purpose: \<short summary for automation registry\>
>
> '   Inputs: \<table/column list\>
>
> '   Outputs: \<tables/columns modified\>
>
> '   Preconditions: \<required conditions\>
>
> '   Postconditions: \<expected resulting state\>
>
> '   Errors: \<categories of errors raised\>
>
> '   Version: \<repeat version\>
>
> '   Author: \<name\>
>
> '   Date: \<YYYY-MM-DD\>
>
> '\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*

This (metadata) information will be automatically recorded by a script that parses the headers of each script in the workbook and published to TBL_AUTO within the Auto tab. A parser may be added in a future release.

Authoritative sources

The authoritative workbook schema is maintained in SCHEMA!TBL_SCHEMA. All validation logic and automation must treat this table as the primary contract.

The Workbook_Schema tab is a human-readable snapshot for review and documentation. It is derived from (and must not contradict) SCHEMA!TBL_SCHEMA.

Any exported schema files (e.g., Workbook_Schema_Long \*.csv) are point-in-time exports for offline review only; they are not authoritative.

Automation Registry refresh (implemented v3.4.0)

1.  Inventory rows are populated by M_Core_Automation (implemented).

2.  scans all Public procedures in StdModules

3.  upserts rows in Auto.TBL_AUTO

4.  overwrites baseline inventory columns: Module, Trigger, Status, Feature, FeatureName, audit fields

5.  flags stale entries

Automation Registry - Feature enrichment (planned / optional)

1.  Enrichment fields are manual for now; a parser may be added later.

2.  parses standardized procedure headers

3.  fills the richer columns (Inputs/Outputs/Helpers/Guards/Tests)

(You are explicitly not locked into implementation yet)

This cleanly explains why many rows currently have blank Inputs/Outputs/etc.

TBL_AUTO is an authoritative inventory (scanner-owned baseline) of current automation

- “Header-parsing enrichment” is planned and not part of v3.4.0

- The old enrichment mechanism (Automation_RegisterFeature) is retired (v3.4.0)

For Workbook/050 Document Version v3.4.0:

- TBL_AUTO is the authoritative inventory of public procedures (scanner-owned baseline columns).

- Enrichment is optional and not required for correctness.

Scanner-owned columns (always overwritten):

Module, Status, Trigger, Feature, FeatureName, Public Entry Point, UpdatedAt, UpdatedBy

(CreatedAt/CreatedBy only set on insert)

User/curation-owned columns (not overwritten by scanner):

Input, Output, Helpers, Errors/Guards, Tests, Notes/Version

The scanner **does not overwrite** Input/Output/Helpers/Errors/Tests/Notes if already populated (curation-preserved).

Table Classes and how they are checked for data integrity by M_Core_HealthCheck:

Input Tables = Validated

Derived Tables = Excluded,

System Tables = excluded from data integrity tests

Each table must have an ActiveRowDriver for driving auditing or active rows only, defined in 055 schema

## Golden Paths & Edge Cases

Golden Path:

- Designers define Suppliers, Components, and BOMs (Top Assemblies);

- PMs add Builds (opening WO’s) for TA’s, generating Demand;

- Buyers add PO Lines for components satisfying Demand;

- Components are issued to Work Orders

- TA’s are shipped, WO’s are closed complete.

User Functions / Automation Modules

Table 1: Current list of modules as of today v3.4.0.

<table>
<colgroup>
<col style="width: 35%" />
<col style="width: 26%" />
<col style="width: 15%" />
<col style="width: 22%" />
</colgroup>
<thead>
<tr>
<th style="text-align: center;">Module Name of .bas file</th>
<th style="text-align: center;"><blockquote>
<p>Module Description</p>
</blockquote></th>
<th style="text-align: center;">Module Status</th>
<th style="text-align: center;">Primary UI Entrypoint</th>
</tr>
</thead>
<tbody>
<tr>
<td colspan="4" style="text-align: center;">Developer-only Modules</td>
</tr>
<tr>
<td style="text-align: center;">Dev_ExportVBAModules</td>
<td style="text-align: center;">Exports all VBA modules to a user-selected folder</td>
<td style="text-align: center;">Implemented v3.3.x</td>
<td style="text-align: center;">UI_</td>
</tr>
<tr>
<td style="text-align: center;">Dev_Generate_ModuleInventory</td>
<td style="text-align: center;">Generates inventory of modules and public procedures</td>
<td style="text-align: center;">Implemented v3.3.x</td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: center;">Dev_Schema_Report_Universal</td>
<td style="text-align: center;">Exports workbook schema (tabs/tables/headers) to CSV/worksheet</td>
<td style="text-align: center;"><p>Implemented</p>
<p>V3.3.x</p></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td colspan="4" style="text-align: center;">User Interaction UI_ Functions</td>
</tr>
<tr>
<td style="text-align: center;">M_Data_Suppliers_Entry</td>
<td style="text-align: center;">Creates new supplier records with validation</td>
<td style="text-align: center;">Implemented 3.4.2</td>
<td style="text-align: center;">UI_New_Supplier</td>
</tr>
<tr>
<td style="text-align: center;">M_Data_Comps_Entry</td>
<td style="text-align: center;"><em>Create new component record</em></td>
<td style="text-align: center;"><em>Implemented 3.4.2</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: center;">M_Core_HealthCheck</td>
<td style="text-align: center;">Runs platform checks (schema + data integrity + gate) and summarizes status</td>
<td style="text-align: center;">Implemented &amp; Active</td>
<td style="text-align: center;">UI_Run_HealthCheck</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Gate</td>
<td style="text-align: center;">Central gate that blocks operational macros unless validators pass</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">N/A</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Automation</td>
<td style="text-align: center;">Authoritative Automation Registry inventory refresh (scan public procedures, upsert TBL_AUTO, flag stale entries)</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">UI_Automation,</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Constants</td>
<td style="text-align: center;">Central constants and identifiers (sheet/table/column names)</td>
<td style="text-align: center;">Implemented &amp; Active</td>
<td style="text-align: center;">N/A</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_DataIntegrity</td>
<td style="text-align: center;">Validates content integrity (required fields, duplicates, etc.)</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">UI_</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Logging</td>
<td style="text-align: center;">Central logging to Log. TBL_LOG</td>
<td style="text-align: center;">Implemented &amp; Active</td>
<td style="text-align: center;">N/A</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Schema</td>
<td style="text-align: center;">Validates workbook structure against TBL_SCHEMA</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">_Schema_Check?</td>
</tr>
<tr>
<td colspan="4" style="text-align: center;"><em>Safety, Checking &amp; Utilities</em></td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Tests</td>
<td style="text-align: center;">Core test harness and structural tests</td>
<td style="text-align: center;">Implemented &amp; Active</td>
<td style="text-align: center;">?</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Toggles</td>
<td style="text-align: center;">Application state toggles (calc/events/screen updating)</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">N/A</td>
</tr>
<tr>
<td style="text-align: center;">M_Core_Utils</td>
<td style="text-align: center;">Shared utilities (table helpers, safe setters, dictionaries)</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">N/A</td>
</tr>
<tr>
<td style="text-align: center;">WorksheetSwitches</td>
<td style="text-align: center;">Navigation helpers</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;">UI_[“To” Worksheet]</td>
</tr>
<tr>
<td style="text-align: center;">Test_Logging</td>
<td style="text-align: center;">Logging test harness</td>
<td style="text-align: center;">Implemented</td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td colspan="4" style="text-align: center;"><em>Future Modules for Planning Purposes</em></td>
</tr>
<tr>
<td style="text-align: right;"><em>M_Inv_Receive</em></td>
<td style="text-align: center;"><em>Create new inventory transaction</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"><em>M_New_PO</em></td>
<td style="text-align: center;"><em>Create New Vendor PO Record</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"><em>M_New_Comp</em></td>
<td style="text-align: center;"><em>Create new component record</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"><em>M_WO_Demand</em></td>
<td style="text-align: center;"><em>Script that calculates demand for all comps</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>New, Edit Supplier</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>New, Edit Component</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>New, Edit Top Assembly (BOM)</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>New, Edit Build (WO)</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>New, Edit PO Line</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>Issue Components to Build</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
<tr>
<td style="text-align: right;"></td>
<td style="text-align: center;"><em>Ship a Build</em></td>
<td style="text-align: center;"><em>FUTURE</em></td>
<td style="text-align: center;"></td>
</tr>
</tbody>
</table>

Edge Cases:

EC-01 Late Supplier (include “late” for status) / Partial Receipts (need option to close line from partial receipt)

EC-02 Missing/Invalid Keys

EC-03 Version Mismatch

EC-04 Duplicate WO / Overlapping Builds (make it easy to omit wo’s from calculations (planning/HOLD/etc))

## Testing Strategy (Starter Kit)

Path Tests: PT-G1 Golden, PT-E1 EC-01, PT-E2 EC-02.

Smoke Tests: Test_DataKeys, Test_ForeignKeys, Test_UniquePNRev, Unique keys, FK resolution, no negative Qty.  
Golden dataset: 10–50 rows per table for regression checks.  
Table-Driven Test Harness (TBL_TEST_CASES table): suite, proc, inputs, assertions, expected outputs.

Semantic versioning (vMAJOR.MINOR.PATCH) logged in ABOUT sheet.  
Macro Archive_Backup creates timestamped backup files.

Note: The TBL_TEST_CASES table is a core component for automated testing, required for the MVP to ensure regression checks.

Per-procedure Test\_() harnesses = developer debugging.

TBL_TEST_CASES = regression suite for ongoing validation.

## Iteration & Deliverables

Each iteration should ship updated .xlsm, Diff note, Short test checklist, Changelog entry and at least one new test case.

## Versioning & Change Control

> Major: add/remove/rename tables or core columns
>
> Minor: add optional columns, add lookup tables
>
> Patch: Notes, descriptions, typo fixes

## Workbook Schema

> *Workbook Schema is defined in 055 – Workbook Schema v3.4.0.xlsx*
>
> *Schema 3.4.0 is the structural baseline; structural changes require a version bump.*
>
> *Schema version in TBL_SCHEMA must match SCHEMA_VERSION in M_Core_Constants.*
>
> *Workbook will **fail validation** otherwise.*

## 

## SemVer in module headers and a small ABOUT cell block.

TBL_CHANGELOG (optional).

Macro signing/trusted locations after stabilization.

Alignment with Governance Docs

> 010 - Documentation Strategy: This doc fulfills the 'Workbook Layout & Functional Strategy' role.  
> 020 - GPT Instructions: Fully aligned on naming, schema validation, error handling, and logging.  
> 030 - Charter: Demand, Inventory, and Users tables now explicitly included.  
> 040 - Vibe: Advanced features (Event Bus, Registry) flagged as extensions, not MVP. Clarity \> Cleverness.
>
> 050 – This Workbook Creation Strategy / Definition Document
>
> 055 – The Schema definition document that defines the workbook layout, tabs, table and column headers

Glossary

BuildID Unique identifier for each Build or Work Order (WO), past examples include NSWO-25-038

Build Name Unique descriptive name of build, past examples include Feasibility Build, pre-DV Build, HF Build.

CompID Unique identifier for each combination of component part number and revision

Description Descriptive name of component, duplicates allowed.

SupplierID Unique identifier for each component supplier

Supplier Name of Supplier

AssemblyID

BomID Unique identifier for each BOM, defined as that which follows BOM\_ on any worksheet tab

OurPN Internal Part Number

OurRev Internal Revision

MOQ1/2/3 Minimum order Quantity one, two and three e.g. MOQ1 would be the smallest order allowed, MOQ2 would be first price break, MOQ3 would be second price break.

POID Unique identifier for each Purchase Order number – one PO ID per PO Number

POLine Unique identifier for every line appearing on every purchase order

QOH Quantity on Hand – physical inventory present, excludes items issued to a build

NetAvailable Quantity available to be issued to a build or builds

Global Note: This workbook will be developed by one developer with the assistance of a custom GPT as identified in 030.

*Inactive suppliers cannot be assigned to components.*

*Only Maintenance can change suppliers from Active → Inactive.*

*Inactivation is forbidden while supplier is referenced in any component.*

**Tab classes**

**Production tabs (user-facing):** contain user workflows, controlled data-entry fields, and reporting.

**System tabs (hidden/protected):** SCHEMA, AUTO, Log, Schema_Check, Data_Check (and similar). These support validation, logging, and automation.

**Developer tabs (hidden/protected, non-shipping optional):** e.g., Dev_ModuleInventory, internal inventories, staging artifacts.

**Visibility policy**

Non-production tabs may be hidden and protected in released versions.

End users are not permitted to modify schema, macros, validation logic, or records outside designated data-entry fields.

Review the following changes as a supervising architect.  
Check for:

1.  Violations of stated schema or naming conventions

2.  Redundant or wrapper-only code

3.  Logic that should live elsewhere

4.  Premature abstraction

5.  Drift from MVP intent

Respond only with:

- KEEP

- MODIFY (with reason)

- DEFER

- REJECT

For Reference and filling in some gaps: Overview of Workbook-User Interaction.

Users begin by creating **Supplier** records, followed by **Component** records. Each component references an existing supplier and is uniquely identified by a PN + Revision combination. Components and suppliers may be inactivated to prevent future use while retaining historical data.

Users then define **Top Assemblies (TAs)** by creating Bills of Material (BOMs). Each TA is represented by a worksheet named BOM\_\<TA Name\> created from a template and contains a table listing required components (PN + Revision) and quantities per assembly. A BOM may include multiple instances of the same PN, at the same or different revisions. Once a TA is referenced by a Work Order, the BOM worksheet name becomes fixed, but BOM contents may still be edited.

Users create **Work Orders (WOs)** by selecting a TA, entering a build quantity and due date. Upon WO creation, and whenever a referenced BOM is edited, the system calculates component demand by PN + Revision based on BOM quantities and WO quantities. Demand is reported both per open WO and as an aggregated total across all open WOs. WOs always reference the current state of a BOM; BOM changes intentionally propagate to existing WOs by updating demand. Allocation priority between WOs is managed manually by the user.

Users review demand and create **Purchase Order (PO) Lines** to satisfy unmet demand. Demand views highlight components where available supply (on-hand plus on-order) does not cover total demand. PO Lines default to ordering the full open demand for a component but may be adjusted by the user. PO issuance is assumed once a PO Line is created.

When components arrive, users record **receipts** against PO Lines. Receipts increase inventory and default to a receiving inspection state. Users then manually disposition received items (e.g., accepted, quarantined, returned to vendor). Inventory is assumed usable unless explicitly dispositioned otherwise. All receipts and disposition changes are tracked in the inventory ledger.

Users **issue inventory** to WOs as a deliberate, manual action. Issuing inventory reduces on-hand and net available quantities and associates the issued components with a specific WO. Inventory is not automatically allocated or consumed.

Users record **shipment of completed TAs** against WOs. Multiple shipments per WO are permitted. WO status is managed manually (e.g., open, complete, on hold, canceled). Recording shipments completes the operational lifecycle.

Throughout the workflow, demand is always derived rather than manually entered, inventory movements are explicit and logged, and automation is designed to provide visibility and traceability while leaving planning, prioritization, and execution decisions under user control.

V3.4.0: Users run UI macros (only UI\_\*). All operational macros are protected by M_Core_Gate. Registry + schema + health checks are the “platform”.

**VBA Code Inventory (Concise)**

Generated: December 22, 2025

Scope: module-level summary of procedures (Subs/Functions) with purpose, inputs, outputs, and assumptions.

## Dev_ExportVBAModules.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Export_All_VBAModules_SaveAsFolder (Sub) | Export all VBA modules to a chosen folder. | VBProject/VBComponents (implicit) | Writes .bas/.cls/.frm(+.frx) files; \_EXPORT_INFO.txt | Requires 'Trust access to VBA project object model' enabled. |
| ConvertUrlToLocalPath (Function) | Convert OneDrive/Office URL paths to local paths when possible. | urlPath As String | String | Assumes local sync/mapping exists. |
| NormalizeFolderPath (Function) | Normalize folder formatting (e.g., trailing slash). | folderPath As String | String | Windows path conventions. |
| ExportVBComponent (Sub) | Export a single VBComponent with correct extension. | vbComp; exportFolder | File created | VBIDE component types are recognized. |
| GetExportExtension (Function) | Map VBComponent type to extension. | vbComp | String | Known VBIDE enums. |
| WriteExportInfo (Sub) | Write export metadata file. | exportFolder; infoText | \_EXPORT_INFO.txt | Folder writable. |
| SafeFileName (Function) | Sanitize names for filesystem. | s As String | String | Windows filename restrictions. |
| FolderExists (Function) | Test for folder existence. | folderPath | Boolean | Filesystem accessible. |
| EnsureFolder (Function) | Create folder if missing. | folderPath | Boolean | Permissions allow creation. |

## Dev_Generate_ModuleInventory.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Generate_ModuleInventory (Sub) | Scan VBProject and generate a module/procedure inventory. | VBProject (implicit) | Inventory report (sheet/file) | Requires VBProject access; parsing is text-based. |
| GetAllModules (Function) | Return list of modules in VBProject. | VBProject (implicit) | Collection/Variant | VBProject accessible. |
| GetModuleCodeLines (Function) | Read code text for a module. | vbComp | String | Module is readable. |
| ParseProcedures (Function) | Parse Sub/Function declarations from code. | moduleText | Procedure metadata collection | Assumes conventional VBA declaration lines. |
| CleanLine (Function) | Normalize/sanitize a line for parsing. | s | String | Text parsing sufficient. |

## Dev_Schema_Report_Universal.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Generate_Workbook_Schema_Report (Sub) | Create schema snapshot of sheets/tables/headers. | Workbook structure (implicit) | Report sheet and/or CSV | ListObjects used for tables. |
| WriteSchemaToSheet (Sub) | Write schema rows to a worksheet. | targetWs; schemaRows | Populated report | Target writable. |
| CollectSchema (Function) | Collect ListObject headers across workbook. | wb | Schema rows collection | Tables are ListObjects. |
| IsWorksheetInScope (Function) | Filter sheets for reporting. | ws | Boolean | Naming conventions define scope. |
| ExportSchemaAsCsv (Sub) | Export schema rows to CSV. | schemaRows; filePath | CSV file written | Filesystem writable. |
| SafeCsv (Function) | Escape/quote values for CSV. | s | String | Standard CSV quoting. |

## M_Core_Automation.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Core_Run_All (Sub) | Run an end-to-end automation sequence (checks + actions). | Workbook state (implicit) | Workbook updated; logs/reports | Depends on gate/schema/integrity modules. |
| Ensure_Audit_Columns (Sub) | Ensure audit fields exist in a ListObject. | lo As ListObject | Columns added if missing | Table structure can be edited. |
| Normalize_Table_Headers (Sub) | Apply standard header formatting/canonicalization. | lo | Header row standardized | Conventions defined and stable. |
| Get_User_Identity (Function) | Resolve user identifier for audit stamping. | Environment (implicit) | String | UserName/env values available. |
| Stamp_Row_Audit (Sub) | Stamp audit values into a row. | lo; rowIndex; isCreate | Row cells updated | Audit columns present/ensured. |
| NowUtcISO (Function) | Return UTC timestamp string. | None | String | System clock correct. |
| Ensure_Table_Exists (Sub) | Confirm a ListObject exists; optionally create/repair. | ws; tableName | Table exists or issue raised/logged | Worksheet exists; writable if creating. |
| Rebuild_Dependent_Validations (Sub) | Refresh data validation that depends on tables. | Workbook state | DV updated | Named ranges/tables consistent. |
| Refresh_All_Pivots (Sub) | Refresh all pivots. | Workbook state | Pivot caches refreshed | Pivot sources valid. |
| Core_Run_Selected (Sub) | Run selected steps based on toggles/arguments. | Step flags (varies) | Partial run results | Toggle model consistent. |

## M_Core_Constants.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| (No procedures) | Central constants for sheet/table/column names and shared tokens. | N/A | N/A | Must be kept aligned with schema and actual headers. |

## M_Core_DataCheck_Audit.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Audit_Check_All (Sub) | Validate audit columns/values across in-scope tables. | Workbook state | Issues logged/reported | Audit standard defined; logging active. |
| Audit_Validate_Table (Function) | Validate one table for audit standard. | lo | Boolean/result | Audit columns have defined names/types. |
| Audit_Report_Issues (Sub) | Emit audit findings. | issues | Report/log output | Destinations available. |

## M_Core_DataIntegrity.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Integrity_Check_All (Sub) | Run full integrity suite (keys/required/relationships). | Workbook state | Issue list + summary | Schema expectations accurate. |
| Check_Table_Exists (Function) | Confirm table existence by sheet/table name. | sheetName; tableName | Boolean | Workbook object model accessible. |
| Check_Required_Columns (Function) | Validate required headers exist. | lo; requiredCols | Boolean/issues | Exact header matching unless normalized. |
| Check_Unique_Key (Function) | Ensure key values unique and nonblank. | lo; keyCol | Boolean/issues | Key column exists. |
| Check_Foreign_Key (Function) | Validate FK values exist in parent PK set. | childLo/childCol; parentLo/parentCol | Boolean/issues | Comparable data types; parent authoritative. |
| Get_Column_Index (Function) | Resolve header to column index. | lo; colName | Long | Exact header match. |
| Get_Column_Values (Function) | Extract column values for validation. | lo; colName | Array/Collection | Empty tables handled. |
| Add_Issue (Sub) | Append standardized issue record. | issues; type; context; detail | issues mutated | Issue record format stable. |
| Is_BlankOrError (Function) | Treat blanks/errors uniformly. | v | Boolean | Excel error variants possible. |
| Check_Required_NonBlank (Function) | Validate required fields not blank. | lo; requiredCols | Boolean/issues | Columns exist. |
| Check_Allowed_Values (Function) | Validate values against allowed set. | lo; colName; allowed | Boolean/issues | Allowed set defined. |
| Build_Lookup_Set (Function) | Build dictionary/set for membership tests. | values | Dictionary-like object | Scripting.Dictionary available (or late-bound). |
| Check_Duplicate_Rows (Function) | Detect duplicates using multiple columns. | lo; cols | Boolean/issues | Concatenation strategy stable. |
| Integrity_Report (Sub) | Emit integrity findings. | issues | Report/log output | Logging/reporting configured. |
| Normalize_Value (Function) | Normalize values for comparison. | v | String/Variant | Normalization acceptable for keys. |
| Check_Date_Logic (Function) | Validate date logic/plausibility. | lo; dateCols | Boolean/issues | Dates are true dates or parseable. |
| Check_Number_Range (Function) | Validate numeric bounds. | lo; colName; min/max | Boolean/issues | Numeric coercion acceptable. |
| Check_Text_Length (Function) | Validate maximum text length. | lo; colName; maxLen | Boolean/issues | String conversion safe. |
| Check_Composite_Key (Function) | Validate composite key uniqueness. | lo; cols | Boolean/issues | Delimiter safe; stable. |
| InScope_Table (Function) | Determine whether a table participates. | ws; lo | Boolean | Scope rules defined. |

## M_Core_Gate.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Gate_CanRun (Function) | Evaluate prerequisites before running automation. | Workbook + environment | Boolean | Logs reasons on failure. |
| Gate_Enforce (Sub) | Stop execution if gate fails; message/log. | None | Execution halted / error/message | User interaction permitted. |
| Gate_Check_VBAProjectAccess (Function) | Verify VBProject access if required. | None | Boolean | Trust Center setting governs access. |
| Gate_Check_SchemaPresent (Function) | Ensure schema assets exist. | None | Boolean | Schema location is fixed. |
| Gate_Check_CoreSheets (Function) | Ensure required tabs exist. | None | Boolean | Sheet names stable. |

## M_Core_HealthCheck.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| HealthCheck_Run (Sub) | Run health suite and generate summary. | Workbook state | Status summary + report/log | Depends on check modules. |
| HealthCheck_Status (Function) | Compute overall status label/code. | issues | String/enum-like | Severity classification exists. |
| HealthCheck_WriteReport (Sub) | Write results to a report sheet/table. | issues; targetWs | Report output | Target exists/writable. |
| HealthCheck_Summarize (Function) | Summarize issues into counts/buckets. | issues | Collection/Dictionary | Issue schema consistent. |
| HealthCheck_ClearPrior (Sub) | Clear prior report outputs. | targetWs | Cleared content | Sheet not protected. |
| HealthCheck_InScope (Function) | Filter tables/sheets for health check. | ws; lo | Boolean | Scope rules defined. |
| HealthCheck_LogBanner (Sub) | Standard banner logging. | runLabel | Log entries | Logging configured. |
| HealthCheck_HasBlocking (Function) | Detect blocking issues. | issues | Boolean | Blocking threshold defined. |

## M_Core_Lockdown.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Lockdown_Apply (Sub) | Apply protection + hide dev artifacts. | Workbook state | Workbook locked down | Passwords/options available; permissions allow. |
| Lockdown_Remove (Sub) | Remove/relax lockdown for development. | Workbook state | Workbook unlocked | Correct password if used. |
| HideDevSheets (Sub) | Hide known dev-only sheets. | Workbook state | Visibility changed | Sheet names match list. |
| ShowDevSheets (Sub) | Unhide dev sheets. | Workbook state | Visibility changed | Same as above. |
| ProtectWorkbookStructure (Sub) | Protect workbook structure. | Optional password | Protection enabled | Excel APIs available. |
| UnprotectWorkbookStructure (Sub) | Unprotect workbook structure. | Optional password | Protection disabled | Correct password if used. |
| ProtectAllSheets (Sub) | Protect worksheets consistently. | Workbook state | Sheets protected | Options align with intended use. |
| UnprotectAllSheets (Sub) | Unprotect worksheets. | Workbook state | Sheets unprotected | Correct password if used. |
| IsDevSheet (Function) | Determine whether sheet is dev-only. | wsName | Boolean | Naming conventions stable. |
| Lockdown_Diagnostics (Sub) | Generate protection/visibility diagnostics. | Workbook state | Diagnostic output | Able to write diagnostics destination. |

## M_Core_Logging.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Log_Info (Sub) | Write informational log record. | msg; context(optional) | Log row appended | Log destination exists/creatable. |
| Log_Warn (Sub) | Write warning log record. | msg; context | Log row appended | Same as above. |
| Log_Error (Sub) | Write error log record (optionally includes Err details). | msg; context; Err (implicit/optional) | Log row appended | Same as above. |

## M_Core_Schema.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Schema_Validate_All (Sub) | Validate workbook against SCHEMA!TBL_SCHEMA. | Workbook state | Pass/fail + issues | Schema is authoritative and correct. |
| Schema_Validate_Table (Function) | Validate one table's headers vs schema. | wsName; tableName | Boolean/issues | Table exists or missing handled as issue. |
| Schema_GetExpectedColumns (Function) | Fetch expected columns from schema. | wsName; tableName | Collection/Variant | Schema table key columns consistent. |
| Schema_GetActualColumns (Function) | Read actual ListObject headers. | lo | Collection/Variant | ListObject has headers. |
| Schema_CompareColumns (Function) | Compare expected vs actual and produce diffs. | expected; actual | Boolean/issues | Comparison rules defined (exact vs normalized). |
| Schema_Report (Sub) | Emit schema validation output. | issues | Report/log output | Destinations available. |
| Schema_TableInScope (Function) | Filter participating tables. | ws; lo | Boolean | Scope rules defined. |

## M_Core_Tests.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Run_Core_Tests (Sub) | Execute core test routines. | Workbook state | Test results/log output | Dev/test mode; safe to mutate workbook. |
| Assert_True (Sub) | Assertion helper. | condition; message | Raises/logs on failure | Failure behavior acceptable. |

## M_Core_Toggles.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Toggle_Get (Function) | Retrieve toggle value by name. | toggleName | Variant/Boolean | Toggles stored in known location. |
| Toggle_Set (Sub) | Persist toggle value by name. | toggleName; value | Toggle updated | Storage writable. |
| Toggle_LoadDefaults (Sub) | Initialize default toggle set. | None | Defaults written | Defaults defined in code/constants. |

## M_Core_Utils.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Nz (Function) | Null/empty coercion helper. | v; fallback | Variant | Consistent blank/error semantics desired. |

## M_Data_Suppliers_Entry.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| NewSupplier (Sub) | Create a new supplier record (UI-driven). | User inputs/selection (implicit) | New row + audit stamps | Supplier table exists; required cols defined. |
| EditSupplier (Sub) | Edit an existing supplier record. | Selected supplier row/ID | Updated row | Row resolvable; keys stable. |
| Supplier_Exists (Function) | Check if supplier exists by key. | supplierKey | Boolean | Key column exists. |
| GetSupplierTable (Function) | Return supplier ListObject. | None | ListObject | Sheet/table names align with constants. |
| Supplier_ValidateRow (Sub) | Validate required fields and formats. | lo; rowIndex | Pass/fail/issues | Required columns known. |
| Supplier_StampAudit (Sub) | Stamp Created/Updated fields. | lo; rowIndex; isCreate | Row updated | Audit columns exist/ensured. |

## M_UI_Validation.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| UI_ShowValidationSummary (Sub) | Display validation results to user. | issues | UI feedback | User interaction permitted. |
| UI_FormatIssue (Function) | Format an issue for display. | issue | String | Issue schema consistent. |
| UI_FocusIssueLocation (Sub) | Navigate user to issue location. | issue | Selection changed | Location resolvable; sheet accessible. |
| UI_NotifyBlocked (Sub) | Notify user of a blocking condition. | message | UI message | UI permitted. |

## temp.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| CreateLockdownDiag (Sub) | Create diagnostics sheet/table (ad hoc). | Workbook state | Diagnostic sheet created | Dev-only; structure changes allowed. |
| Temp_Run (Sub) | Temporary runner for ad-hoc tests. | None | Varies | Dev-only. |

## Test_Logging.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| TestLogging (Sub) | Logging smoke test. | None | Log entries created | Log destination exists/creatable. |

## WorksheetSwitches.bas

| **Procedure** | **Purpose** | **Inputs** | **Outputs** | **Assumptions / Notes** |
|----|----|----|----|----|
| Switches_Initialize (Sub) | Initialize worksheet switch states. | Workbook state | Switch values set | Switch storage exists. |
| Switch_Get (Function) | Get switch value. | switchName | Variant/Boolean | Switch defined. |
| Switch_Set (Sub) | Set switch value. | switchName; value | Switch updated | Storage writable. |
| Switch_EnableAll (Sub) | Enable all switches (dev convenience). | None | Switches set | Safe to alter globally. |
| Switch_DisableAll (Sub) | Disable all switches. | None | Switches set | Same. |
| ApplySwitchesToSheets (Sub) | Apply switch config to sheet behaviors. | Workbook state | Properties/behaviors adjusted | Expected sheets exist. |
