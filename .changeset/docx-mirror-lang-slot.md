---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Screen readers now hear each DOCX text run in the language of its script: Latin text uses the run's Latin language instead of its East Asian one, so English in documents with a Chinese or Japanese East Asian default is no longer announced as Chinese or Japanese.
