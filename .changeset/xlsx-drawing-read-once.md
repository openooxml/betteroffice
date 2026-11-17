---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
---

A drawing that a sheet names with extra leading slashes, or through several relationships, is read once, so its charts and drawing objects are no longer listed several times, and such a workbook no longer fails to open with too many charts.
