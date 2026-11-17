---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

Breaking: `DrawCmd` gains `Image`, `build_display_list_with_geometry` and `build_print_display_list_with_charts` take pictures, and `WorkbookHandle` requires `mediaBytes`. Screen and print display lists carry worksheet pictures, drawn by the native PNG export and by `paintDisplayList` given `resolveImage`; the JavaScript PNG export, bundled editor and viewer do not paint them yet.
