---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

`repackDocx` and Rust `Document::save` return the original package for an unchanged model, updating only core properties when stamping. `repackDocxWithWarnings` names the parts rewritten after a model edit.
