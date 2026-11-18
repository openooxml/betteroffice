---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A field that spans paragraphs no longer repeats the later paragraphs inside itself. Its structured code and result give their count in the new optional `blockCount` instead, so fields left open across many paragraphs no longer duplicate entries or repeat work.
