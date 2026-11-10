---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

Fixes collaborative documents that stopped laying out for every collaborator after an undo removed a table cell another collaborator's table still showed. The cell now reads as empty, and typing into it recreates it.
