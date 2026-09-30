---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Opening a document no longer waits for the East Asian or complex-script fonts its text names but never uses, so the first page no longer waits on a large CJK font for Latin-only text.
