---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Editing a table header cell now renames the table column as Excel does, so saved workbooks open in Excel without repair; duplicate or cleared headers become unique names such as `Name2` or `Column1`.
