---
"@betteroffice/rust-crates": patch
---

Adds `parse_docx_s9_wire_with_media_table`, `parse_docx_s9_preview_with_media_table` and `media_table_parts`, which parse without inflating image parts, plus `RetainedPackage` and `MediaTable`. Existing entry points are unchanged.
