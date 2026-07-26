---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Saving a workbook now preserves the parts the model does not represent — charts, drawings, pivot tables, comments, macros, custom XML and their relationships — instead of rebuilding the package and dropping them. Chartsheets keep their type, rich shared strings keep per-run formatting, defined names follow sheet renames and structural edits, and an edited save drops the stale calculation chain so Excel recalculates on open.
