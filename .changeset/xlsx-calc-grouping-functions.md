---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

The calculation engine now evaluates GROUPBY, PIVOTBY, TRIMRANGE and EUROCONVERT and accepts function names such as SUM as arguments. PRODUCT of a block without numbers returns 0 and ARRAYTOTEXT keeps 15 significant digits, as in Excel. GETPIVOTDATA and FORECAST.ETS keep Excel's saved values.
