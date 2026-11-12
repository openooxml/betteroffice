---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Cell fonts with `<u val="single"/>`, `double` or accounting underlines now read as underlined. Double kinds paint two lines and accounting kinds span the cell below the descenders in every renderer. Text commands add `underlineStyle`, captured fonts add `underline_style`; Rust `Font.underline` is now a `RunUnderline`.
