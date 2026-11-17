---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Fixes collaborative table editing where splitting a cell again after a collaborator removed the previous split cell made the new cell invisible, and where redoing a shared table deletion left nested cells visible.
