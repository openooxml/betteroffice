---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A comment with ID 0 can now be re-anchored with `setCommentRanges` and keeps its range when the document is saved after an edit elsewhere. Comment, bookmark, move-range and note markers without a `w:id` are now ignored instead of being read as ID 0.
