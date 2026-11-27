---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Text composed with an input method is now typed into the document as it is composed, in the document's font, instead of drawn in an overlay. `insertText` takes an optional `replaceBefore`, the text to replace right before a collapsed caret.
