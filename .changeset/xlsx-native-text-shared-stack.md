---
'@betteroffice/rust-crates': minor
'@betteroffice/xlsx-native': patch
---

Native XLSX rendering shapes and draws cell text through the shared font store and glyph cache, in about half the time. A no-break hyphen in a font without one draws as a hyphen, not a missing-glyph box. `FontStore` gains `register_shared`.
