---
"@betteroffice/docx-react": minor
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Plugin `geometry.getAnchorGeometry`, `readAnchorGeometry` and `readAnchorGeometries` now resolve targets in headers and footers, with rects on every page that paints them. Unbuilt display pages name their header and footer parts in `hfParts` (`HfParts` in Rust). On by default.
