---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Proposing and applying edits reuse one projection of each story until the document changes, so a batch of edits to a long document no longer rebuilds the body for every edit.
