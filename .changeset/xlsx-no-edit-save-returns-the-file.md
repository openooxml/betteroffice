---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A save with no edit since open, showing the active sheet the file stores, returns the file unchanged when its archive can be reused as read, also when the open recalculated a file asking for a full calculation on load. `recalculate()` before saving still writes its values.
