---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": minor
---

Ctrl+End reaches the used-range corner, while other keyboard navigation reaches the sheet bounds. Distant selections remain visible after navigation and zooming. Selection queries accept `usedRange` to request used bounds without computing a region. Rust: adds `SelectionDirection::UsedRange`; `SelectionDirection` is `#[non_exhaustive]`.
