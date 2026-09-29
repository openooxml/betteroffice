---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Opening a DOCX now surfaces non-fatal warnings for parts without content types and duplicate drawing, bookmark, and Word paragraph IDs. Hosts addressing paragraphs by raw Word IDs should prefer session anchors with `resolveParagraphAnchor()` or call `persistParagraphIds()` to repair duplicates.
