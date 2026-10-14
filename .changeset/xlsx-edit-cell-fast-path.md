---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Single-cell edits apply directly instead of rebuilding the workbook's editing state, so cell edits on large workbooks are much faster, with identical results.
