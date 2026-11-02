---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

Formulas using functions that cannot be computed offline (CUBE functions, WEBSERVICE, RTD, STOCKHISTORY, IMAGE, add-in calls and similar), and formulas reading them, keep the value Excel saved, even inside IFERROR, and the calculation report lists them as `unevaluatedCells` in JS and `unevaluated` in Python.
