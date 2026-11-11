---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Cell fonts with `<u val="single"/>`, `double` or accounting underlines now read as underlined, and double kinds paint two lines. Text commands add `underlineStyle` and captured fonts add `underline_style`; `Font.underline` in Rust is now a `RunUnderline`.
