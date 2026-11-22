---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A save with no edit and no change of the active sheet since open returns the file unchanged, also when the open recalculated a file asking for a full calculation on load. `recalculate()` before saving still writes its values.
