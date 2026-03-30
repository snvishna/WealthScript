---
description: The official CI/CD pipeline for this workspace. This workflow safely validates, documents, and commits code to version control.
---

# Execution Pipeline

You must execute the following steps sequentially. If any step fails, **ABORT immediately** and report to the user. Do not proceed until the failure is resolved.

---

## Step 1 — Test Generation
Execute `@.agents/workflows/generate-tests.md` to ensure all new logic has acceptance tests.

> [!IMPORTANT]
> **TDD Requirement:** If adding new behavior, the test must be written BEFORE implementation and must describe *expected behavior*, not the implementation. A test that only asserts a formula string is correct is insufficient — it must assert what the formula *does*.

---

## Step 2 — Integration & E2E Execution (HARD GATE)
Run `runAllTests()` from the Google Apps Script editor (or via clasp run).

- The suite spins up a live sandbox spreadsheet (`Integration.test.gs`).
- **ABORT if ANY assertion fails.** Fix the logic, restart the pipeline.
- E2E tests must exercise live data paths — **do not mock volatile dependencies** (e.g., GOOGLEFINANCE, Drive REST API). If a test needs to inject data, it must still allow the formula chain to run natively (e.g., `SpreadsheetApp.flush()` + `Utilities.sleep()`).

---

## Step 3 — Mandatory Visual Smoke Test (NON-SKIPPABLE)
Run `runFirstTimeSetup()` on a live test spreadsheet and visually verify each area changed by the current diff:

| Check | Criterion |
|---|---|
| **Dashboard Current Value** | Brokerage rows show live sums from Holdings tab (not 0) |
| **Conditional Formatting** | Only brokerage rows are muted/italic; others are default |
| **Brokerage Holdings** | GOOGLEFINANCE prices populated; Total Value = Qty × Price |
| **Snapshots** | `captureSnapshot()` inserts a row with correct Net Worth |
| **Settings** | All config sections at correct rows; hyperlinks resolve |
| **New feature area** | Verify the specific behavior you just implemented works end-to-end |

> [!CAUTION]
> Do not skip this step because the automated tests passed. Tests pass against static fixtures; visual smoke tests catch formatting, UX, and real-data integration failures. The SUMIF→0 and all-blue CF bugs would both have been caught in 30 seconds of visual inspection.

---

## Step 4 — Pre-Flight Lint
Execute `@.agents/workflows/lint-code.md`. Ensure SOLID principles and GAS coding standards are met.

---

## Step 5 — Documentation Sync
Execute `@.agents/workflows/sync-readme.md`. Ensure `README.md` reflects all changes including Settings tab row references.

---

## Step 6 — Build Artifacts
```bash
./scripts/build.sh
```
Compiles all `src/` modules into `deploy/code.gs` and copies `appsscript.json` to `deploy/`.

---

## Step 7 — Semantic Commit
- Analyze the diff of all staged and unstaged files.
- Generate a Conventional Commits formatted message (e.g., `fix(dashboard): replace SUMIF with SUMPRODUCT for cross-tab volatile formula support`).
- `git add . && git commit -m "<message>"`

---

## Step 8 — Confirmation
Output a success matrix confirming: tests passed, visual smoke clear, lint passed, docs synced, artifact built, commit hash.