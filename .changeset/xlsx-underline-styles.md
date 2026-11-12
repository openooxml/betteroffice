---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Cell fonts with `<u val="single"/>`, `double` or accounting underlines now read as underlined, and double kinds paint two lines. Accounting underlines draw as plain single or double for now. Text commands add `underlineStyle`, captured fonts add `underline_style`; Rust `Font.underline` is now a `RunUnderline`.
