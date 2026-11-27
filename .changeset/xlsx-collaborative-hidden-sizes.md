---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Collaborative spreadsheet sessions keep the authored height or width of hidden rows and columns: saving writes it instead of zero, and unhiding restores it on every replica.
