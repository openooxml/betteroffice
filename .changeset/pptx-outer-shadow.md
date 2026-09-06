---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Parse `a:effectLst/a:outerShdw` and paint the drop shadow of filled and outlined shapes in Canvas and PNG exports.

Preserve shadow scale and alignment across schema-18 snapshots, following the schema-17 bitmap-effect migration. Bound per-slide shadow work in Canvas and PNG exports.
