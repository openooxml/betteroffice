---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": minor
---

Fixes Ctrl/Meta+Arrow data edges, two-stage Ctrl/Meta+A selection, and navigation through hidden tracks, enabled by default in editors. `selectionKeyReducer` accepts optional `SelectionData` and `SelectionHidden`; `moveFocus` accepts `hidden`; `GridMeta` includes optional `hiddenRowSpans` and `hiddenColSpans`. Rust: `SelectionDirection` is `#[non_exhaustive]`.
