---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
---

A table or comments part that a sheet relates several ways is read once, so tables are not listed twice, comment counts are not multiplied, and such a workbook no longer fails to open with too many tables.
