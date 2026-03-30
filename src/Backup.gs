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
function forceBackup(ss_inject) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
  backupToGitHub(ss, false);
  backupToGoogleDrive(ss, false);
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
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet (for injection)
 * @param {boolean} [silent=false] - If true, suppresses UI alerts on success.
 * @returns {boolean} Whether the backup was attempted and succeeded.
 */
function backupToGitHub(ss_inject, silent = false) {
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();
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
 * Pure helper: returns the OAuth token for Drive REST API calls.
 * @returns {string}
 */
function _getDriveToken() {
  return ScriptApp.getOAuthToken();
}

/**
 * Pure helper: Creates a folder via Drive REST API v3 (drive.file scope).
 * @param {string} folderName
 * @returns {string} The created folder ID
 */
function _createDriveFolder(folderName) {
  const token = _getDriveToken();
  const meta = { name: folderName, mimeType: "application/vnd.google-apps.folder" };
  const resp = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(meta),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error("Drive folder creation failed: " + resp.getContentText());
  return JSON.parse(resp.getContentText()).id;
}

/**
 * Pure helper: Creates a plain-text file inside a Drive folder via REST API.
 * @param {string} folderId - Parent folder ID
 * @param {string} fileName
 * @param {string} content - File text content
 * @returns {string} The created file ID
 */
function _createDriveFile(folderId, fileName, content) {
  const token = _getDriveToken();
  const boundary = "wealthscript_boundary";
  const body =
    "--" + boundary + "\r\n" +
    "Content-Type: application/json\r\n\r\n" +
    JSON.stringify({ name: fileName, parents: [folderId] }) + "\r\n" +
    "--" + boundary + "\r\n" +
    "Content-Type: text/plain\r\n\r\n" +
    content + "\r\n" +
    "--" + boundary + "--";
  const resp = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "multipart/related; boundary=" + boundary
    },
    payload: body,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error("Drive file creation failed: " + resp.getContentText());
  return JSON.parse(resp.getContentText()).id;
}

/**
 * Pure helper: Gets the web view link for a Drive file/folder by ID.
 * @param {string} fileId
 * @returns {string} The webViewLink
 */
function _getDriveFileLink(fileId) {
  const token = _getDriveToken();
  const resp = UrlFetchApp.fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?fields=webViewLink`,
    { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) return `https://drive.google.com/drive/folders/${fileId}`;
  return JSON.parse(resp.getContentText()).webViewLink || `https://drive.google.com/drive/folders/${fileId}`;
}

/**
 * Google Drive Backup: serializes the enriched ledger to a dated JSON file.
 * Uses Drive REST API v3 with drive.file scope — only accesses app-created files.
 * @param {SpreadsheetApp.Spreadsheet} [ss_inject] - Target spreadsheet (for injection)
 * @param {boolean} [silent=false] - Suppresses UI alerts on success.
 * @param {string} [folderId_inject] - Optional pre-created folder ID (for E2E tests)
 * @returns {{success: boolean, folderId: string}} Result object
 */
function backupToGoogleDrive(ss_inject, silent = false, folderId_inject = null) {
  const FOLDER_NAME = "WealthScript \u2014 Backups";
  const ss = ss_inject || SpreadsheetApp.getActiveSpreadsheet();

  try {
    const backupData = _buildEnrichedBackup(ss);
    const jsonContent = JSON.stringify(backupData, null, 2);

    const folderId = folderId_inject || _createDriveFolder(FOLDER_NAME);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH-mm");
    const fileName = `net_worth_${timestamp}.json`;
    _createDriveFile(folderId, fileName, jsonContent);

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `\u2705 Google Drive Backup Successful!\n\nFolder: "${FOLDER_NAME}"\nFile: ${fileName}`
      );
    }
    return { success: true, folderId };
  } catch (e) {
    Logger.log("Google Drive backup error: " + e.message);
    if (!silent) SpreadsheetApp.getUi().alert("\u274c Drive Backup Failed:\n" + e.message);
    return { success: false, folderId: null };
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
