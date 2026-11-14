---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/docx-native': patch
---

Formatting a paragraph mark inside a block content control, such as a form field, no longer rewrites the rest of the document on save: tracked changes, proofing marks, drawings and fields elsewhere are kept as they were.
