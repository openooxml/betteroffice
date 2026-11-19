---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

A save with created list definitions no longer fails or writes a wrong package when relationships or numbering use another prefix or quoting, keeps list ids clear of ids the document names, and writes only definitions in use. New: the edit crate's `list_definition_keys_in_use`.
