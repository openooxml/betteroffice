---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A comment with ID 0 can now be re-anchored with `setCommentRanges` and keeps its range when saved after an edit elsewhere. Comment ranges, bookmarks and notes without a `w:id` are now ignored instead of being read as ID 0.
