---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A document whose chart part cannot be read now opens with that chart missing, instead of the whole file being refused; paragraphs, tables and text come through intact. Chart parts also get their own parse budget, so a valid document whose charts carry large cached series opens with every chart, and the part the parser declined is still written back untouched on save.
