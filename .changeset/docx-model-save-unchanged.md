---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

For unchanged models, `repackDocx` saves byte for byte with `updateModifiedDate: false` and no modifier change, and Rust `Document::save` preserves the original package except for core-property stamping. `repackDocxWithWarnings` lists changed package members when returning rewritten bytes.
