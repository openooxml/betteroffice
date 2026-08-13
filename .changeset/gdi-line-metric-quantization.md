---
"@betteroffice/rust-crates": patch
---

Reject non-finite line-metric sizes and cap line boxes at Word's 1638pt limit. GDI quantization and version-4 typographic metrics ship as off-by-default experiments; measured against Word 16.112, both are worse than the float default.
