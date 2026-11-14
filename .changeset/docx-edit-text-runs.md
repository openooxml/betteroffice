---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
---

`applyEdits` text steps and new paragraphs accept `runs` with bold, italic, underline, strike, colour, highlight, size, font and superscript/subscript marks instead of plain `text`. Rust `EditOperation` and `ParagraphInput` gain a `runs` field.
