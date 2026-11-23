---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A save with no edit since open returns the file unchanged, also after an open that recalculated a full-calculation file, while the stored active sheet is shown and the archive is reusable. Applying a collaborative update or `recalculate()` counts as an edit.
