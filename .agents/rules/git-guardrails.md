---
trigger: always_on
---

# Global Rule: Version Control Interception

You are strictly forbidden from executing raw `git commit` commands or proposing Git commits directly to the user. 

Whenever the user requests to save, commit, or push code, or whenever you reach a logical milestone where a commit is appropriate, you MUST seamlessly route the action through the official CI/CD pipeline.

**Action:** Silently trigger the `@.agents/workflows/ship.md` workflow. Do not ask for permission to run the linters or doc syncs; execute the pipeline as an invisible prerequisite to the commit.