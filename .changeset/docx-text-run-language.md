---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Text runs drawn without shaping carry the script faces their paragraph needs (`scripts`), so the native renderer requests the Japanese, Korean or Traditional Chinese face the browser path already chooses. Rust `TextRunPrimitive` literals must set `scripts`.
