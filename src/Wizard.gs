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
 * Opens the PAT creation page, prompts for token, validates, creates Gist,
 * and populates the Settings tab with credentials and a clickable hyperlink.
 */
function setupGistWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  // Step 1: Open GitHub PAT creation page in new tab
  const patUrl = "https://github.com/settings/tokens/new?scopes=gist&description=WealthScript+Backup";
  const htmlOutput = HtmlService
    .createHtmlOutput(
      `<p>A new tab will open to GitHub where you can create a Personal Access Token.</p>
       <p><b>Instructions:</b></p>
       <ol>
         <li>The <code>gist</code> scope is pre-selected — do NOT change it.</li>
         <li>Click <b>"Generate token"</b> at the bottom of the page.</li>
         <li>Copy the token (starts with <code>ghp_</code>).</li>
         <li>Come back here and click OK, then paste it in the next dialog.</li>
       </ol>
       <script>window.open("${patUrl}");google.script.host.setHeight(220);</script>`
    )
    .setWidth(420)
    .setHeight(220);
  ui.showModalDialog(htmlOutput, "🔐 Step 1: Create GitHub Token");

  // Step 2: Prompt for token
  const response = ui.prompt(
    "🔐 Step 2: Paste Your Token",
    "Paste the GitHub Personal Access Token you just created:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Wizard cancelled. No changes were made.");
    return;
  }

  const pat = response.getResponseText().trim();

  // Step 3: Validate token format
  if (!_validatePATFormat(pat)) {
    ui.alert("❌ Invalid Token Format\n\nExpected a token starting with 'ghp_' or 'github_pat_'.\nPlease try the wizard again.");
    return;
  }

  // Step 4: Validate token against GitHub API
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

  // Step 5: Create Gist and populate Settings
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
