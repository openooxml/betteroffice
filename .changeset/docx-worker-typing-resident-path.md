---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Typing in the worker-only docx editor lays out each text command through the resident incremental path instead of a full pass. Suggesting edits, empty insertions and sessions whose resident input is not ready keep the generic edit and full pass.
