---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Conditional formatting rules that apply a differential format paint their fill, font color, bold and italic on the cells they match, on screen and in PNG renders. On by default. Rust `xlsx_render` gains `CellFormats`.
