---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": minor
---

Fixes Ctrl/Meta+Arrow data edges, two-stage Ctrl/Meta+A selection and navigation through hidden rows and columns in editors, on by default. `selectionKeyReducer`, `moveFocus` and `GridMeta` take optional selection data and hidden spans, and Rust `SelectionDirection` is `#[non_exhaustive]`.
