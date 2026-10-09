---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Scrolling large workbooks repaints faster: the grid's row and column geometry is reused between frames until the workbook changes.
