---
'@betteroffice/docx': minor
'@betteroffice/rust-crates': minor
'@betteroffice/docx-native': patch
---

Formatting selections spanning tables (select-all then bold) formats cell text. Paragraphs that would lose content on save stay unchanged. Edits and saves report which and why. Have `YrsSession` hosts and mocks return `YrsFormatReceipt` from `toggleMark`, `formatRange` and `clearFormatting`, and initialize new `S13ParagraphIds` and `S13SplicedPart` fields.
