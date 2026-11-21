---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Typing in a table cell of a long document no longer rebuilds the body's paragraph and segment indexes on every keystroke: a text edit that changed only one story keeps the other stories' indexes.
