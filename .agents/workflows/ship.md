---
description: The official CI/CD pipeline for this workspace. This workflow safely validates, documents, and commits code to version control.
---

# Execution Pipeline
You must execute the following steps sequentially. If any step fails, ABORT the pipeline immediately and report the failure to the user. Do not proceed to the next step until the failure is resolved.

1. **Test Generation:**
   - Execute `@.agents/workflows/generate-tests.md` to ensure all new logic has coverage.

2. **Test Execution:**
   - Simulate a run of the newly generated test functions. If any `Assert` throws an error, ABORT the pipeline, fix the underlying logic in `code.gs`, and restart the pipeline.
   
3. **Pre-Flight Check (Linting):**
   - Execute the logic defined in `@.agents/workflows/lint-code.md`.
   - Ensure all SOLID principles and Google Apps Script standards are met.
   
4. **Documentation Sync:**
   - Execute the logic defined in `@.agents/workflows/sync-readme.md`.
   - Ensure `README.md` reflects the absolute latest state of the codebase, including any new configurations required in the Settings tab.

5. **Semantic Versioning & Commit:**
   - Analyze the diff of all staged and unstaged files.
   - Generate a professional, Conventional Commits formatted commit message (e.g., `feat(tracker): add AI insights pipeline`).
   - Execute `git add .` followed by `git commit -m "<your generated message>"`.

6. **Confirmation:**
   - Output a success matrix to the chat, confirming the linting passed, docs were updated, and the hash of the new commit.
