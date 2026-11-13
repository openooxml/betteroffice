---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

DOCX tables now paint the cell shading and borders of their table style's conditional regions, such as header rows, first columns and banded rows, for the regions the table's `tblLook` turns on. On by default.
