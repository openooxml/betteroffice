---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Parse runs, paragraphs, tables, rows and cells that Word wraps in `w:customXml`, `w:smartTag` or a row/cell `w:sdt`; their text was dropped on open and therefore on save.
