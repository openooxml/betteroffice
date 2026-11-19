---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Conditional formats paint fill, font color, bold and italic, on screen and in PNG renders, only where the result is certain; after a cell edit, value paste or recalculation, an undecided cell keeps its previous format. On by default. Rust `xlsx_render` gains `CellFormats` and `build_display_list_with_geometry_and_formats`.
