---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A full recalculation reuses the formulas the previous dependency graph already parsed instead of parsing every formula again. Over 208 public workbooks the total recalculation time fell by 12 to 14 percent.
