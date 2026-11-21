---
'@betteroffice/docx': minor
'@betteroffice/rust-crates': minor
'@betteroffice/docx-native': patch
---

Formatting a selection that spans a table, such as select all then bold, now formats the text in its cells. Paragraphs a save would lose content from stay unchanged, and edits and saves report which and why. `YrsSession` hosts and mocks must return `YrsFormatReceipt` from `toggleMark`, `formatRange` and `clearFormatting`, and Rust literals must set `S13ParagraphIds.whole_parts`, `S13ParagraphIds.touched_parts` and `S13SplicedPart.format_only`.
