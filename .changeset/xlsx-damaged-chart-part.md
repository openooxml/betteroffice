---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A workbook whose chart or drawing part cannot be read now opens with that chart missing, instead of the whole file being refused; cells, formulas and styles come through intact. The part it declined to read is still written back untouched on save.
