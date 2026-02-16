# AGENTS.md

## Purpose
These instructions define how Codex should propose and execute code-change plans in this repository.

## Default planning behavior
When proposing work, always provide a structured task list using this template for each task:

- **Task**
- **Why**
- **Depends on**
- **Risk level** (`low` / `medium` / `high`)
- **How to verify**
- **Can skip?** (`yes/no` + consequence)

## Execution guidance
- Treat tasks as a plan, not a batch script.
- Default to running tasks **one at a time**.
- Default to running tasks **in order** unless explicitly marked independent.
- Clearly label dependencies and prerequisites.
- If a task is skipped, explicitly state downstream impact before continuing.

## Validation expectations
- After each meaningful task, run a targeted validation step (tests, lint, build, or focused checks).
- Prefer incremental verification over end-only verification.

## Change management
- Prefer small, logically scoped commits.
- Keep high-risk changes isolated from low-risk refactors when possible.

## Communication style
- Be concise but explicit about assumptions, dependencies, and risk.
- If uncertainty exists, present recommended order and a safe fallback path.
