---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": minor
---

Fixes Ctrl/Meta+Arrow data edges, two-stage Ctrl/Meta+A selection and navigation through hidden rows and columns in editors, on by default. `selectionKeyReducer` takes optional `SelectionData` and `SelectionHidden`, `moveFocus` takes `hidden`, `GridMeta` has optional `hiddenRowSpans` and `hiddenColSpans`, and Rust `SelectionDirection` is `#[non_exhaustive]`.
