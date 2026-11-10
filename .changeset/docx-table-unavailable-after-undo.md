---
'@betteroffice/rust-crates': patch
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

Fixes collaborative documents that stopped laying out after two collaborators undid edits to one table in crossing order. That table is now skipped; table edits and saves refuse with a coded error until it is restored by redo or deleted.
