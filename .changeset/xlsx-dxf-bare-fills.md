---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

A conditional format's fill without `patternType`, as Excel writes its presets, keeps its color; a conditional fill naming two different colors paints none. Rust `DifferentialFormat` gains `is_released_reading_of`.
