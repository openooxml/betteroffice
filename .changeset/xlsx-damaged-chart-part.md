---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A workbook whose chart or drawing part cannot be read now opens with that chart missing, instead of the whole file being refused; cells, formulas and styles come through intact, and the part it declined to read is written back untouched on save. Such a workbook is cell-editable but structurally frozen: inserting or deleting rows and columns, and renaming or removing a sheet, are refused, because nothing can move what a part it cannot read names — the deal a chart part no sheet anchors already got. A chart part a drawing names but the package does not hold now freezes those edits too, where before it let them strand the chart silently.
