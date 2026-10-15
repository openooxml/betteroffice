---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Formulas referencing another workbook, using R1C1 `INDIRECT` or adding to date text keep their saved results after an edit instead of `#REF!` or `#VALUE!`. A range used as a single value reads the cell on the formula's row or column, as in Excel.
