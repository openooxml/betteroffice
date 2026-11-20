---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Paragraph styles whose `numPr` has no nonzero `numId`, such as Word's Subtitle, now apply without being mistaken for numbered styles. Truly numbered styles remain unsupported by the edit batch.
