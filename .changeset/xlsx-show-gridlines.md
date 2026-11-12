---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Sheets whose first sheet view sets `showGridLines="0"` render on screen without gridlines, as in Excel, and keep the setting on save. Print gridlines still follow the `gridlines` argument. Rust `SheetFormat` gains `hide_gridlines`.
