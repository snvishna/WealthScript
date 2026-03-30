# ADR-001: Google Drive Backup — DriveApp vs. Drive REST API

**Date:** 2026-03-30  
**Status:** Accepted  
**Deciders:** @snvishna

---

## Context

WealthScript backs up the user's financial ledger to Google Drive as a dated JSON file on every snapshot. Two implementation approaches were evaluated:

### Option A: Drive REST API (`UrlFetchApp` to `googleapis.com/drive/v3`)
- **Scope required:** `drive.file` — app can only see files it created itself, not the rest of the user's Drive.
- **Security posture:** Excellent. True minimal-access principle. Even if WealthScript code were compromised, it could not read any pre-existing user files.
- **Setup required:**
  1. User must manually apply `appsscript.json` manifest (enable in Project Settings, paste content).
  2. User must enable the **Google Drive API** in Google Cloud Console > APIs & Services.
- **Technical complexity:** Higher. Requires multipart upload form construction, OAuth token retrieval, and manual HTTP response parsing.

### Option B: DriveApp (native GAS SDK)
- **Scope required:** `drive` — full Drive read/write access.
- **Security posture:** Acceptable. GAS apps are sandboxed to the authorizing user's context; full Drive scope is the standard for any GAS app using DriveApp.
- **Setup required:** None. DriveApp auto-enables on first authorization. Zero Cloud Console interaction.
- **Technical complexity:** Low. Two lines: `DriveApp.getFoldersByName()`, `folder.createFile()`.

---

## Decision

**Chosen: Option B — DriveApp.**

The security gain from `drive.file` does not justify the UX cost for this tool's target user. WealthScript is a personal finance tracker used by individual non-technical users. Requiring them to navigate Google Cloud Console before the Drive backup feature works is a significant friction point that will cause setup abandonment.

The `drive` scope is the standard for any Google Apps Script application using file operations. It is scoped to the authorizing user's own account — there is no third party, no server, no data exfiltration risk distinct from any other GAS app.

---

## Consequences

- **Positive:** Setup is 4 steps: paste code, save, refresh, authorize. No Cloud Console.
- **Positive:** No `SERVICE_DISABLED` 403 errors during onboarding.
- **Negative:** OAuth consent screen shows `"See, edit, create and delete files in Google Drive"` — broader wording than `drive.file` would have shown. This is documented in the README OAuth Scopes section with a plain-English explanation.
- **Neutral:** `appsscript.json` is still maintained to lock down the other 4 scopes (preventing GAS auto-detection from requesting broader permissions on non-Drive services).

---

## Revisit Criteria

Reconsider the Drive REST API approach if:
- Google provides a way to auto-enable the Drive API without Cloud Console interaction.
- WealthScript is published as an official Google Workspace Add-on (in which case the manifest and API enablement are handled by the publishing process).
- A technical path is found to use `drive.file` scope without requiring users to manually apply the manifest and enable APIs.
