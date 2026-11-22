---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

A save with no edit since open returns the file unchanged, also when the open recalculated a file asking for a full calculation on load. `recalculate()` before saving still writes its values. Supersedes `xlsx-save-recalculated-values` on the open case.
