---
description: Analyzes the current state of the Apps Script codebase and synchronizes the README.md to reflect all active features, triggers, and configurations.
---

1. **Analyze Codebase:** Read the latest state of `src/**/*.gs`. Identify all active functions, external API calls (`UrlFetchApp`), and Google Sheets UI modifications.
2. **Map Configurations:** Identify all required variables mapped to the `Settings & Config` tab (e.g., API Keys, Target Net Worth, Annual Return rates).
3. **Draft Updates:** Compare the newly discovered architecture against the existing `README.md`.
4. **Patch:** Silently patch the `README.md` with the new updates. Ensure formatting is clean, professional, and uses Markdown tables for configuration variables.