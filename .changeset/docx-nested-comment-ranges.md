---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Comments anchored inside a table cell or a block content control now keep their ranges when the document is saved after an edit elsewhere, and can be re-anchored with `setCommentRanges`.
