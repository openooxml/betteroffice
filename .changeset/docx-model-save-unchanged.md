---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

`repackDocx` and Rust `Document::save` return the original bytes when the model and written core properties are unchanged. `repackDocxWithWarnings` reports rewritten package members when their names are available.
