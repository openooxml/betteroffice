---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A table cell whose own font has no color, in a workbook whose normal font has one, keeps its automatic text color instead of taking the table style's, as in Excel. A header with its own light fill no longer shows white text.
