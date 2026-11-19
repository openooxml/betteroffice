---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/rust-crates": minor
---

Text inside `w:dir` and `w:bdo` is shown, found and kept on save and edit; runs carry `directionScopes` (`DirectionScope`). Rust: `Run` gains the public field `direction_scopes`, so struct literals listing every field must add it.
