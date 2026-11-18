---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Conditional formatting rules paint their fill, font color, bold and italic on matching cells, on screen and in PNG renders, only where the result is certain; otherwise a cell paints none. On by default. Rust `xlsx_render` gains `CellFormats` and `build_display_list_with_geometry_and_formats`.
