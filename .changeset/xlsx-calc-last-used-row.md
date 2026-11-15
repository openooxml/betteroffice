---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Formulas that need a sheet's last used row, such as SUMPRODUCT and whole-column references, no longer rescan every cell of the sheet to find it.
