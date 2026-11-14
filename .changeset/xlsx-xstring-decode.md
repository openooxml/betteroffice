---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Cell text decodes Excel's `_xHHHH_` escapes, so `_x000D_` reads as a carriage return, and saves write control characters the way Excel does. Opened formulas keep control characters; new ones store them as `CHAR(n)`, and chart caches over such cells save.
