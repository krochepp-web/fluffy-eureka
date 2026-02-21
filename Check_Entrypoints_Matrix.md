# Canonical Check Entrypoints Matrix

The schema definition tab (`SCHEMA` / `TBL_SCHEMA`) is the single source of truth whenever a checker validates workbook schema and/or required column headers.

| Check | Canonical entrypoint(s) | Purpose | Primary outputs | Intended trigger points |
|---|---|---|---|---|
| Schema check | `RunSchemaCheck` -> `Schema_Check` | Validate workbook tabs/tables/headers against `SCHEMA!TBL_SCHEMA`. | `Schema_Check` sheet with issue rows (row 2+). | Before releases, after schema edits, before gate decisions. |
| Data check | `RunDataCheck` -> `Data_Check` | Validate required, unique, keys, and FK integrity using schema rules. | `Data_Check` sheet with issue rows (row 2+). | Before releases, after data-rule edits, before gate decisions. |
| Gate check | `RunGateCheck` | Strict pass/fail decision using schema + data checker outputs. | Boolean decision (`True`/`False`), plus optional logging and user summary. | Any operation that must block when workbook is not ready. |
| Diagnostics | `RunDiagnostics` | Optional comprehensive report that runs schema/data checks and summarizes status (including optional gate run if available). | Summary message/log plus refreshed `Schema_Check` and `Data_Check` sheets. | Manual troubleshooting, developer validation, pre-release review. |

## Wrapper compatibility and deprecation

- `UI_Run_GateCheck`: aligned to strict gate decision (`RunGateCheck`).
- `UI_Run_AllChecks`: compatibility alias to `UI_Run_GateCheck` semantics.
- `UI_Run_HealthCheck`: compatibility alias to `RunDiagnostics`.
- `ValidateSchema`: compatibility shim only (deprecated); use `RunSchemaCheck` / `Schema_Check`.
