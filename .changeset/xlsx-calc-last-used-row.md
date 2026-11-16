---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Recalculation is faster for formulas that depend on a sheet's last used row, such as SUMPRODUCT over whole columns.
