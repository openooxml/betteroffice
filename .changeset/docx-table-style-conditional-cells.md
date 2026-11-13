---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

DOCX tables now show the cell shading and borders of their table style's conditional regions, such as header rows, first columns and banded rows, where the table's `tblLook` turns them on. Saving keeps these as style formatting rather than writing them into the cells.
