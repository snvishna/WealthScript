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
 * Uses a SINGLE modal dialog with inline link + embedded token input
 * to avoid the dual-popup overlap issue.
 */
function setupGistWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");

  if (!configSheet) {
    ui.alert("⚠️ Settings tab not found.\n\nPlease run '🚀 Run First Time Setup' first.");
    return;
  }

  const patUrl = "https://github.com/settings/tokens/new?scopes=gist&description=WealthScript+Backup";
  const htmlContent = `
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; color: #1a1a1a; margin: 0; }
      .step { margin-bottom: 10px; display: flex; align-items: flex-start; gap: 10px; }
      .step-num { flex-shrink: 0; background: #2563EB; color: white; border-radius: 50%;
                  width: 24px; height: 24px; text-align: center; line-height: 24px; font-size: 13px; }
      a.btn { display: inline-block; background: #2563EB; color: white !important; padding: 10px 20px;
              border-radius: 6px; text-decoration: none; font-weight: bold; margin: 8px 0 12px; }
      a.btn:hover { background: #1d4ed8; }
      code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
      input[type=text] { width: 100%; padding: 10px; border: 2px solid #e2e8f0; border-radius: 6px;
                          font-size: 14px; font-family: monospace; box-sizing: border-box; margin: 6px 0; }
      input[type=text]:focus { outline: none; border-color: #2563EB; }
      .submit-btn { background: #059669; color: white; border: none; padding: 10px 24px;
                     border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 4px; }
      .submit-btn:hover { background: #047857; }
      .submit-btn:disabled { background: #9ca3af; cursor: not-allowed; }
      .status { margin-top: 10px; padding: 10px; border-radius: 6px; font-size: 13px; display: none; }
      .status.error { background: #fef2f2; color: #be123c; display: block; }
      .status.success { background: #ecfdf5; color: #059669; display: block; }
      .status.loading { background: #eff6ff; color: #2563EB; display: block; }
      hr { border: none; border-top: 1px solid #e2e8f0; margin: 12px 0; }
    </style>

    <div class="step"><span class="step-num">1</span><span>Click the button below to open GitHub's token page:</span></div>
    <a class="btn" href="${patUrl}" target="_blank">🔐 Open GitHub Token Page →</a>
    <div class="step"><span class="step-num">2</span><span>The <code>gist</code> scope is pre-selected — just click <b>"Generate token"</b></span></div>
    <div class="step"><span class="step-num">3</span><span>Copy the token and paste it below:</span></div>
    <hr>
    <input type="text" id="tokenInput" placeholder="ghp_xxxxxxxxxxxxxxxxxxxx" />
    <button class="submit-btn" id="submitBtn" onclick="submitToken()">✅ Connect to GitHub</button>
    <div class="status" id="statusMsg"></div>

    <script>
      function submitToken() {
        var token = document.getElementById('tokenInput').value.trim();
        var btn = document.getElementById('submitBtn');
        var status = document.getElementById('statusMsg');
        
        if (!token) {
          status.className = 'status error';
          status.textContent = 'Please paste your token above.';
          return;
        }
        
        btn.disabled = true;
        btn.textContent = '⏳ Validating...';
        status.className = 'status loading';
        status.textContent = 'Validating token and creating your private Gist...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              status.className = 'status success';
              status.textContent = '✅ ' + result.message;
              btn.textContent = '✅ Done!';
              setTimeout(function() { google.script.host.close(); }, 2500);
            } else {
              status.className = 'status error';
              status.textContent = '❌ ' + result.message;
              btn.disabled = false;
              btn.textContent = '✅ Connect to GitHub';
            }
          })
          .withFailureHandler(function(err) {
            status.className = 'status error';
            status.textContent = '❌ Error: ' + err.message;
            btn.disabled = false;
            btn.textContent = '✅ Connect to GitHub';
          })
          ._processGistToken(token);
      }
    </script>`;

  const htmlOutput = HtmlService
    .createHtmlOutput(htmlContent)
    .setWidth(480)
    .setHeight(380);
  ui.showModalDialog(htmlOutput, "🔐 Setup GitHub Backup");
}

/**
 * Server-side handler called from the wizard HTML dialog.
 * Validates the token, creates a Gist, and populates Settings.
 * @param {string} token - The raw PAT string from the dialog input.
 * @returns {{success: boolean, message: string}}
 */
function _processGistToken(token) {
  const pat = (token || "").trim();

  if (!_validatePATFormat(pat)) {
    return { success: false, message: "Invalid format. Expected a token starting with 'ghp_' or 'github_pat_'." };
  }

  // Validate against GitHub API
  try {
    const testResponse = UrlFetchApp.fetch("https://api.github.com/user", {
      headers: { "Authorization": "Bearer " + pat, "Accept": "application/vnd.github.v3+json" },
      muteHttpExceptions: true
    });
    if (testResponse.getResponseCode() !== 200) {
      return { success: false, message: "GitHub rejected this token. Ensure it has 'gist' scope." };
    }
  } catch (e) {
    return { success: false, message: "Network error: " + e.message };
  }

  // Create Gist
  const gistId = autoCreateGist(pat);
  if (!gistId) {
    return { success: false, message: "Token is valid but Gist creation failed. Check Apps Script logs." };
  }

  // Populate Settings tab
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Settings & Config");
  if (configSheet) {
    configSheet.getRange("B13").setValue(pat);
    configSheet.getRange("B14").setValue(gistId);

    const gistUrl = _buildGistUrl(gistId);
    const richGistLink = SpreadsheetApp.newRichTextValue()
      .setText(gistUrl)
      .setLinkUrl(gistUrl)
      .build();
    configSheet.getRange("B15").setRichTextValue(richGistLink);
  }

  return { success: true, message: `Connected! Gist ID: ${gistId}. Every snapshot will now auto-sync to GitHub.` };
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
    const folderId = _createDriveFolder(FOLDER_NAME);
    const folderUrl = _getDriveFileLink(folderId);

    const richDriveLink = SpreadsheetApp.newRichTextValue()
      .setText(folderUrl)
      .setLinkUrl(folderUrl)
      .build();
    configSheet.getRange("B16").setRichTextValue(richDriveLink);

    ui.alert(`✅ Google Drive Backup Connected!\n\nFolder: "${FOLDER_NAME}"\n\nA clickable link has been added to your Settings tab.\nEvery snapshot will now auto-sync a dated JSON file here.`);
  } catch (e) {
    ui.alert("❌ Drive Setup Failed:\n" + e.message);
  }
}
