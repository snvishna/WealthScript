/**
 * ==========================================
 * BACKUP SETUP WIZARD
 * Guided flows for GitHub Gist & Google Drive
 * ==========================================
 */

/**
 * Pure helper: validates that a string looks like a GitHub PAT.
 * Accepts classic tokens (ghp_*) and fine-grained tokens (github_pat_*).
 * @param {string} token - The raw token string from user input.
 * @returns {boolean}
 */
function _validatePATFormat(token) {
  if (!token || typeof token !== 'string') return false;
  const trimmed = token.trim();
  return /^ghp_[A-Za-z0-9]{36,}$/.test(trimmed) || /^github_pat_[A-Za-z0-9_]{20,}$/.test(trimmed);
}

/**
 * Pure helper: builds Gist URL from an ID.
 * @param {string} gistId
 * @returns {string}
 */
function _buildGistUrl(gistId) {
  return `https://gist.github.com/${gistId}`;
}

/**
 * Guided GitHub Gist Backup Wizard.
 * Uses an inline clickable link (popup-blocker safe) instead of window.open().
 * Validates token, creates Gist, populates Settings with credentials + hyperlink.
 */
function setupGistWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  // Step 1: Show instructions with a clickable link (no popup blocker issues)
  const patUrl = "https://github.com/settings/tokens/new?scopes=gist&description=WealthScript+Backup";
  const htmlOutput = HtmlService
    .createHtmlOutput(
      `<style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; color: #1a1a1a; }
        .step { margin-bottom: 12px; }
        .step-num { display: inline-block; background: #2563EB; color: white; border-radius: 50%;
                    width: 24px; height: 24px; text-align: center; line-height: 24px; font-size: 13px; margin-right: 8px; }
        a.btn { display: inline-block; background: #2563EB; color: white !important; padding: 10px 20px;
                border-radius: 6px; text-decoration: none; font-weight: bold; margin: 12px 0; }
        a.btn:hover { background: #1d4ed8; }
        code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
      </style>
      <div class="step"><span class="step-num">1</span> Click the button below to open GitHub's token page:</div>
      <a class="btn" href="${patUrl}" target="_blank">🔐 Open GitHub Token Page →</a>
      <div class="step"><span class="step-num">2</span> The <code>gist</code> scope is pre-selected — just click <b>"Generate token"</b></div>
      <div class="step"><span class="step-num">3</span> Copy the token (starts with <code>ghp_</code>)</div>
      <div class="step"><span class="step-num">4</span> Close this dialog, then paste it in the next prompt</div>`
    )
    .setWidth(460)
    .setHeight(280);
  ui.showModalDialog(htmlOutput, "🔐 Step 1 of 2: Create GitHub Token");

  // Step 2: Prompt for token
  const response = ui.prompt(
    "🔐 Step 2 of 2: Paste Your Token",
    "Paste the GitHub Personal Access Token you just created (starts with ghp_):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Wizard cancelled. No changes were made.");
    return;
  }

  const pat = response.getResponseText().trim();

  if (!_validatePATFormat(pat)) {
    ui.alert("❌ Invalid Token Format\n\nExpected a token starting with 'ghp_' or 'github_pat_'.\nPlease try the wizard again from the WealthScript menu.");
    return;
  }

  // Step 3: Validate token against GitHub API
  try {
    const testResponse = UrlFetchApp.fetch("https://api.github.com/user", {
      headers: { "Authorization": "Bearer " + pat, "Accept": "application/vnd.github.v3+json" },
      muteHttpExceptions: true
    });
    if (testResponse.getResponseCode() !== 200) {
      ui.alert("❌ Token Validation Failed\n\nGitHub rejected the token. Please ensure it has 'gist' scope and try again.");
      return;
    }
  } catch (e) {
    ui.alert("❌ Network Error\n\n" + e.message);
    return;
  }

  // Step 4: Create Gist and populate Settings
  const gistId = autoCreateGist(pat);
  if (!gistId) {
    ui.alert("❌ Gist Creation Failed\n\nThe token is valid but Gist creation failed. Check Apps Script logs for details.");
    return;
  }

  configSheet.getRange("B6").setValue(pat);
  configSheet.getRange("B7").setValue(gistId);

  const gistUrl = _buildGistUrl(gistId);
  const richGistLink = SpreadsheetApp.newRichTextValue()
    .setText(gistUrl)
    .setLinkUrl(gistUrl)
    .build();
  configSheet.getRange("B8").setRichTextValue(richGistLink);

  ui.alert(`✅ GitHub Backup Connected!\n\nYour private Gist has been created and linked.\nGist ID: ${gistId}\n\nEvery snapshot will now auto-sync to GitHub.`);
}

/**
 * Guided Google Drive Backup Setup.
 * Creates the backup folder (if needed), retrieves its URL,
 * and populates the Settings tab with a clickable hyperlink.
 */
function setupDriveBackup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");
  const FOLDER_NAME = "WealthScript \u2014 Backups";

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  try {
    const folderIterator = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(FOLDER_NAME);
    const folderUrl = folder.getUrl();

    const richDriveLink = SpreadsheetApp.newRichTextValue()
      .setText(folderUrl)
      .setLinkUrl(folderUrl)
      .build();
    configSheet.getRange("B9").setRichTextValue(richDriveLink);

    ui.alert(`✅ Google Drive Backup Connected!\n\nFolder: "${FOLDER_NAME}"\n\nA clickable link has been added to your Settings tab.\nEvery snapshot will now auto-sync a dated JSON file here.`);
  } catch (e) {
    ui.alert("❌ Drive Setup Failed:\n" + e.message);
  }
}
