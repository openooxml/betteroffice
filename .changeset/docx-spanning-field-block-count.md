---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

A field that spans paragraphs no longer repeats the later paragraphs inside itself. Its structured code and result give their count in the new optional `blockCount` instead, and 0 means they are no longer tracked.
