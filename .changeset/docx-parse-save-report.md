---
'@betteroffice/rust-crates': minor
---

`docx-parse` adds `write_docx_s13_parts_with_report`, which returns the parts writer's save diagnostics as typed values beside the bytes. `KeptParagraph`, `DroppedParagraph` and `DroppedStructure` become public, and `KeptParagraph.scope` is now a `String`.
