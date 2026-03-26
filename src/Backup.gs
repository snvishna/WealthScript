/**
 * Builds an enriched backup payload including accounts, dashboard KPIs, and latest snapshot.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {Object} Complete backup payload
 */
function _buildEnrichedBackup(ss) {
  const mainSheet = ss.getSheetByName("Dashboard & Ledger");
  const snapSheet = ss.getSheetByName("Snapshots");

  // Account-level data
  const dataRange = mainSheet.getRange("A7:K80").getValues();
  const accounts = _buildLedgerSnapshot(dataRange);

  // Dashboard KPI summary
  const summary = {
    netWorthUSD:   mainSheet.getRange("B2").getValue() || 0,
    grossWorthUSD: mainSheet.getRange("B3").getValue() || 0,
    netWorthSecondary1:  mainSheet.getRange("E2").getValue() || 0,
    grossWorthSecondary1: mainSheet.getRange("E3").getValue() || 0,
    netWorthSecondary2:  mainSheet.getRange("H2").getValue() || 0,
    grossWorthSecondary2: mainSheet.getRange("H3").getValue() || 0,
    liquidNetWorthUSD: mainSheet.getRange("B4").getValue() || 0,
    lockedNetWorthUSD: mainSheet.getRange("E4").getValue() || 0,
    fireProgress:      mainSheet.getRange("I4").getValue() || 0,
  };

  // Latest snapshot row (if exists)
  let latestSnapshot = null;
  if (snapSheet && snapSheet.getLastRow() > 1) {
    const snapRow = snapSheet.getRange(2, 1, 1, 13).getValues()[0];
    latestSnapshot = {
      date:         snapRow[0],
      netUSD:       snapRow[1],
      liquidUSD:    snapRow[2],
      lockedUSD:    snapRow[3],
      grossUSD:     snapRow[4],
      valueDelta:   snapRow[8],
      pctGrowth:    snapRow[9],
      fireProgress: snapRow[10],
      autoInsight:  snapRow[11],
      manualNotes:  snapRow[12]
    };
  }

  return {
    snapshotDate: new Date().toISOString(),
    spreadsheetId: ss.getId(),
    summary,
    latestSnapshot,
    accounts
  };
}

/** Manual trigger: runs both Gist and Drive backups with UI alerts. */
function forceBackup() {
  backupToGitHub(false);
  backupToGoogleDrive(false);
}

/**
 * Pure helper: transforms a raw 2D ledger data range into a structured JSON array.
 * Used by both backupToGitHub() and backupToGoogleDrive().
 * @param {Array<Array<*>>} dataRange - 2D array from sheet.getRange().getValues()
 * @returns {Array<Object>} Structured account objects
 */
function _buildLedgerSnapshot(dataRange) {
  const snapshot = [];
  for (let i = 0; i < dataRange.length; i++) {
    const account = String(dataRange[i][0]);
    if (!account) continue;
    snapshot.push({
      Account:      account,
      AssetClass:   String(dataRange[i][1]),
      Currency:     String(dataRange[i][2]),
      InitialCapital: Number(dataRange[i][3]) || 0,
      CurrentValue:   Number(dataRange[i][4]) || 0,
      ExchangeRate:   Number(dataRange[i][5]) || 0,
      TaxRate:        Number(dataRange[i][6]) || 0,
      GrossWorthUSD:  Number(dataRange[i][7]) || 0,
      NetWorthUSD:    Number(dataRange[i][8]) || 0,
      Status:   String(dataRange[i][9]),
      Remarks:  String(dataRange[i][10])
    });
  }
  return snapshot;
}

/**
 * Checks if GitHub Gist backup is configured in Settings.
 * @param {SpreadsheetApp.Sheet} configSheet
 * @returns {boolean}
 */
function _isGistConfigured(configSheet) {
  if (!configSheet) return false;
  const pat = configSheet.getRange("B13").getValue();
  const gistId = configSheet.getRange("B14").getValue();
  return pat && pat !== "PASTE_GITHUB_TOKEN_HERE" && gistId && gistId !== "PASTE_GIST_ID_HERE";
}

/**
 * Disaster Recovery: Serializes live ledger into enriched JSON and pushes to a private GitHub Gist.
 * Silently skips if not configured (no errors thrown).
 * @param {boolean} silent - If true, suppresses UI alerts on success.
 * @returns {boolean} Whether the backup was attempted and succeeded.
 */
function backupToGitHub(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!_isGistConfigured(configSheet)) {
    // Silently skip — not configured
    return false;
  }

  const githubToken = configSheet.getRange("B13").getValue();
  const gistId = configSheet.getRange("B14").getValue();
  const backupData = _buildEnrichedBackup(ss);

  const payload = {
    "description": "WealthScript Automated Backup",
    "files": {
      "net_worth_backup.json": {
        "content": JSON.stringify(backupData, null, 2)
      }
    }
  };

  const options = {
    "method": "PATCH",
    "headers": {
      "Authorization": "Bearer " + githubToken,
      "Accept": "application/vnd.github.v3+json",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.github.com/gists/" + gistId, options);
    if (response.getResponseCode() === 200) {
      if(!silent) SpreadsheetApp.getUi().alert("✅ GitHub Backup Successful!\n\nYour enriched ledger data has been securely versioned in your private GitHub Gist.");
      return true;
    } else {
      if(!silent) SpreadsheetApp.getUi().alert("❌ GitHub API Error:\n" + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log("GitHub backup error: " + e.message);
    if(!silent) SpreadsheetApp.getUi().alert("❌ Script crashed:\n" + e.message);
    return false;
  }
}

/**
 * Google Drive Backup: serializes the enriched ledger to a dated JSON file.
 * Creates a "WealthScript — Backups" folder in Drive automatically.
 * Silently skips if Drive access fails.
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 * @returns {boolean} Whether the backup succeeded.
 */
function backupToGoogleDrive(silent = false) {
  const FOLDER_NAME = "WealthScript \u2014 Backups";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const backupData = _buildEnrichedBackup(ss);
    const jsonContent = JSON.stringify(backupData, null, 2);

    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(FOLDER_NAME);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
    const fileName = `net_worth_${timestamp}.json`;
    folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `\u2705 Google Drive Backup Successful!\n\nFolder: "${FOLDER_NAME}"\nFile: ${fileName}`
      );
    }
    return true;
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
    return false;
  }
}

/**
 * Helper: Creates a Secret Gist via GitHub API
 */
function autoCreateGist(pat) {
  const payload = {
    "description": "WealthScript Automated Backup",
    "public": false,
    "files": { "net_worth_backup.json": { "content": "{\n  \"status\": \"Initialized\"\n}" } }
  };
  const options = {
    "method": "POST",
    "headers": {
      "Authorization": "Bearer " + pat,
      "Accept": "application/vnd.github.v3+json",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.github.com/gists", options);
    if (response.getResponseCode() === 201) {
      return JSON.parse(response.getContentText()).id;
    } else {
      Logger.log("Failed to create Gist: " + response.getContentText());
      return "";
    }
  } catch (e) {
    Logger.log("Error creating Gist: " + e.message);
    return "";
  }
}
