---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Screen readers hear each DOCX text run in its script's language: Latin text uses the run's Latin language, not its East Asian one, so English in documents with a Chinese or Japanese East Asian default is no longer announced as Chinese or Japanese.
