---
description: Analyzes newly written or modified functions in `src/**/*.gs` and generates isolated, idempotent unit tests in `tests/`.
---

# Execution Steps
1. **Analyze Code Delta:** Review the most recent additions or modifications to `src/**/*.gs`.
2. **Identify Test Candidates:** Focus strictly on pure business logic, mathematical projections (e.g., the FIRE timeline calculations), and data parsing. 
3. **Mock External Services:** - You MUST NOT write tests that make live calls to `UrlFetchApp` (OpenAI, RapidAPI) or `MailApp`.
   - Use dependency injection or mock response objects to simulate API success/failure states.
4. **Mock Google Sheets Interactions:**
   - For functions generating UI or reading Sheets, mock the `SpreadsheetApp` responses using static 2D arrays so the tests can run without modifying the actual live Net Worth Tracker.
5. **Draft & Patch:** Generate the test functions using the `Assert` library and append them to the appropriate file in `tests/`. Output a brief summary of the test coverage to the chat.