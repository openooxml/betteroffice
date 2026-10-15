---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Saving no longer writes inherited formatting into the document: run, cell and table properties are saved relative to their styles, and table row edits leave other rows unchanged. Toggle properties such as bold now combine across styles as OOXML specifies, which can change rendering.
