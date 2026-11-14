---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Excel tables render their table style: header and total rows, banded rows and columns, and first and last columns take the fill, font colour and bold of a built-in or workbook-defined style. Rust `Table` gains `style`.
