---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

A conditional format's fill without `patternType`, as Excel writes its presets, keeps its color, and a solid one reads `bgColor`. Rust `DifferentialFormat::released_reading` becomes `is_released_reading_of`.
