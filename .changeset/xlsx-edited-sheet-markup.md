---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Keep unmodeled row, column, and cell markup on edited sheets. A save now
patches only the cells, rows, and columns an edit changed instead of
reserializing the sheet, so the rest of the sheet survives byte for byte.
