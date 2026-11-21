---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": minor
---

The first picture bullet in a document's numbering now takes up its space at the start of the first paragraph, as in Word, including when that paragraph is in a table. Rust literals must set `NumberingDefinitions.first_picture_bullet` and `ParagraphAttrs.leading_inline_box`.
