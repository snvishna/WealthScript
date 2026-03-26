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
 * Disaster Recovery: Serializes live ledger into JSON and pushes to a private GitHub Gist.
 * @param {boolean} silent - If true, suppresses UI alerts on success (used for chained snapshotting)
 */
function backupToGitHub(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    if(!silent) SpreadsheetApp.getUi().alert("Settings tab missing. Please run First Time Setup.");
    return;
  }

  const githubToken = configSheet.getRange("B6").getValue();
  const gistId = configSheet.getRange("B7").getValue();

  if (!githubToken || githubToken === "PASTE_GITHUB_TOKEN_HERE" || !gistId || gistId === "PASTE_GIST_ID_HERE") {
    if(!silent) SpreadsheetApp.getUi().alert("Disaster Recovery not configured.\n\nPlease add your GitHub PAT and Gist ID in the Settings tab.");
    return;
  }

  const dataRange = sheet.getRange("A7:K80").getValues();
  const backupData = _buildLedgerSnapshot(dataRange);

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
      if(!silent) SpreadsheetApp.getUi().alert("✅ Backup Successful!\n\nYour live JSON data has been securely versioned in your private GitHub Gist.");
    } else {
      if(!silent) SpreadsheetApp.getUi().alert("❌ GitHub API Error:\n" + response.getContentText());
    }
  } catch (e) {
    if(!silent) SpreadsheetApp.getUi().alert("❌ Script crashed:\n" + e.message);
  }
}

/**
 * Google Drive Backup: serializes the live ledger to a dated JSON file.
 * Creates a "WealthScript — Backups" folder in Drive automatically.
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 */
function backupToGoogleDrive(silent = false) {
  const MAX_DRIVE_BACKUPS = 24; // ~2 years of monthly snapshots
  const FOLDER_NAME = "WealthScript \u2014 Backups";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard & Ledger");
  if (!sheet) return;

  try {
    const dataRange = sheet.getRange("A7:K80").getValues();
    const accounts = _buildLedgerSnapshot(dataRange);
    const jsonContent = JSON.stringify({
      snapshotDate: new Date().toISOString(),
      spreadsheetId: ss.getId(),
      accounts
    }, null, 2);

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
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
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
