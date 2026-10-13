---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Typing in the worker-only docx editor lays out text commands through the resident incremental path instead of a full pass, paints the typed text from the caret's line before the worker replies and then reconciles to the reply, and applies worker frames to the retained display pages in place, so only pages changed in place render again.
