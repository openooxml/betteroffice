---
'@betteroffice/rust-crates': patch
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
---

Fixes collaborative documents that stopped laying out after an undo removed a table cell another collaborator's table still showed. That cell now reads as empty and refuses edits until the undo is redone or its row is deleted.
