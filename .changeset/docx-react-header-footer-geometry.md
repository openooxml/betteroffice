---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Plugin `getAnchorGeometry` now resolves targets in headers and footers, including the default worker viewers, with `rects` on every built page that paints them. Unbuilt display pages name their header and footer parts in `hfParts` (`HfParts` in Rust).
