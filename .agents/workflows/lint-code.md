---
description: Performs a semantic, architectural code review on the current state of `src/**/*.gs` against our global coding standards.
---

# Context
Reference the established engineering principles here: `@.agents/rules/apps-script-standards.md`

# Steps
1. **Analyze:** Read the latest version of all files in `src/`.
2. **Evaluate Architecture:** - Check for global scope pollution (are there loose functions that should be in a Class/IIFE?).
   - Verify Separation of Concerns (is UI logic separated from Data fetching?).
3. **Evaluate SOLID Principles:**
   - Identify any function violating the Single Responsibility Principle.
   - Check for hardcoded credentials instead of dependency injection from the Settings tab.
4. **Evaluate Quality & Performance:**
   - Ensure all new primary functions and classes have strict JSDoc typing (`@param`, `@returns`).
   - Flag any instance of `.getValue()` or `.setValue()` inside a loop.
5. **Output the Report:**
   - Generate a concise Code Review summary in the chat. 
   - Use a pass/fail matrix. 
   - If there are failures, generate an Artifact with the proposed refactored code to fix the violations. Do NOT silently overwrite the file.