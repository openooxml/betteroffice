---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Saving a workbook no longer writes a second copy of cells kept inside `mc:AlternateContent` in a worksheet; an edit that would change such a cell is refused with an error before anything changes.
