---
description: Analyzes the current state of the Apps Script codebase and synchronizes the README.md to reflect all active features, triggers, and configurations.
---

## Step 1 — Analyze Codebase
Read the latest state of `src/**/*.gs`. Identify all active functions, external API calls (`UrlFetchApp`), and Google Sheets UI modifications.

## Step 2 — Deployment File Audit (REQUIRED)
List every file in the `deploy/` directory:
```bash
ls deploy/
```
**For each file listed, verify it is explicitly mentioned in the README Quick Start Guide.**

- `deploy/code.gs` → must appear in Phase 1 with copy-paste instructions
- `deploy/appsscript.json` → must appear in Phase 1 as a required manifest step
- Any future `deploy/*.json` config files → must have a corresponding setup step

> If any `deploy/` file is NOT mentioned in the Quick Start Guide, STOP and add it before proceeding. This is the most common source of missing documentation.

## Step 3 — Settings Reference Audit
Scan `src/Builders.gs` for all `setValue()` and `setValues()` calls that write default content to the `Settings & Config` tab. Cross-reference against the **Settings & Config — Full Reference** table in the README. Update any rows that have shifted, been added, or been renamed.

## Step 4 — Feature & Trigger Audit
Identify all:
- Custom menu items (from `onOpen()` in `Menu.gs`) — verify each is described in the README
- Time-based triggers registered in `runFirstTimeSetup()` — verify schedule and handler are documented
- New `src/*.gs` modules added since last sync — verify the File Structure table is up to date

## Step 5 — OAuth Scope Audit
Read `deploy/appsscript.json`. Verify:
- Every scope listed there has a row in the **🔐 OAuth Scopes & Privacy** table in the README
- No scope has been added to the code without being added to the manifest and the README table

## Step 6 — Patch
Silently patch `README.md` with discovered updates. Ensure formatting is clean, professional, and uses Markdown tables for all configuration and scope references.