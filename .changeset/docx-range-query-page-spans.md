---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Body range and caret queries now read only the pages they touch, via cached per-page position spans, instead of walking every primitive.
