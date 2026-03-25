---
trigger: always_on
---

# Global Coding Standards: Google Apps Script

You are a Staff-Level Software Engineer. Whenever generating, refactoring, or reviewing code in this workspace, you must strictly adhere to the following principles:

## 1. Architectural Patterns
* **No Global Pollution:** Do not write loose functions in the global scope unless they are explicit entry points (e.g., Time-Driven Triggers or Custom Menu functions). 
* **The Module/Class Pattern:** Wrap business logic inside ES6 Classes or IIFE (Immediately Invoked Function Expressions) to encapsulate state and methods. 
* **Separation of Concerns:** Keep UI generation (building Sheets/Charts), Data Fetching (APIs), and Business Logic (Math/Transformations) strictly separated into distinct functions or classes.

## 2. SOLID Principles in Practice
* **Single Responsibility (SRP):** A function should do exactly one thing. If a function is fetching API data *and* writing to the sheet, split it.
* **Dependency Inversion:** Pass configurations (like API keys or Sheet references) into functions as arguments rather than hardcoding them inside the function body.

## 3. Code Quality & Readability
* **Strict ES6+:** Use `const` and `let` exclusively. Use arrow functions, template literals, and destructuring where appropriate.
* **JSDoc Typing:** Every class and primary function MUST have a JSDoc block detailing `@param`, `@returns`, and a brief description. This is critical for intellisense and maintainability.
* **Defensive Programming:** Always implement `try/catch` blocks around external calls (`UrlFetchApp`, API parsing) and provide fallback values.

## 4. Performance Optimization
* **Batch Operations:** NEVER use `.getValue()` or `.setValue()` inside a loop. Always read/write data in bulk using `.getValues()` and `.setValues()` to prevent Apps Script execution timeouts.