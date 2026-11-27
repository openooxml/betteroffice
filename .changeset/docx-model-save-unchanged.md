---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

`repackDocx` and Rust `Document::save` return the original package for an unchanged model when the save leaves its core properties, content types and relationships as they were. `repackDocxWithWarnings` names the parts rewritten after a model edit.
