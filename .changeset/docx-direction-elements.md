---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/rust-crates": minor
---

Text inside `w:dir` and `w:bdo` is shown, found and kept on save and editor edits; Rust `Document` single-run text replacement drops the scopes. TypeScript: runs carry optional `directionScopes`. Rust: `Run` gains public `direction_scopes`, so exhaustive struct literals must add it; `DirectionScope` and `DirectionKind` are exported.
