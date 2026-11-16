---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

Breaking: Rust `DrawCmd` gains `Image`, `build_display_list_with_geometry` takes pictures, and `WorkbookHandle` requires `mediaBytes`. Worksheet pictures draw in PNG export and in `paintDisplayList` given `resolveImage`; the bundled editor and viewer do not paint them yet.
