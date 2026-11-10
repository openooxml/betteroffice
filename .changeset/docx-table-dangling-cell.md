---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Fixes collaborative table editing where undoing a split after another collaborator changed the same table, or splitting again after they removed the new cell, left every cell lookup failing until a reload.
