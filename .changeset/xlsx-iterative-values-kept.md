---
'@betteroffice/rust-crates': minor
---

Workbooks with iterative calculation keep the circular values Excel saved instead of saving 0, and those cells are reported as not recalculated. Adds the public field `iterate` to `xlsx_model::Workbook`, so struct literals must set it.
