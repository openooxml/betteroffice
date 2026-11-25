---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": patch
---

`repackDocx` and `repackDocxWithWarnings` preserve untouched DOCX parts, and `createDocx` preserves the source package when given an `originalBuffer`. Native saves and Rust and Python document saves also preserve untouched parts when saving model edits.
