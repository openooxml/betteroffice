---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

Restore inherited column formatting on undo and redo of structural edits. History retains the original ordered column-style runs, including overlapping, clipped and entirely deleted runs, so restored columns and sheets render with their original formatting.
