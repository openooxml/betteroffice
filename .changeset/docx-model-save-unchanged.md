---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

`repackDocx` and Rust `Document::save` return an unchanged model as the original package, updating only the modified date when stamping. `repackDocxWithWarnings` and `save_report_with_warnings` name the parts rewritten after a model edit.
