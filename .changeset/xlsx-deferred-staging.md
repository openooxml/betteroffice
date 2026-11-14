---
'@betteroffice/rust-crates': minor
'@betteroffice/xlsx': patch
'@betteroffice/python-xlsx': patch
---

Editing a standalone XLSX workbook no longer builds its collaboration state: edits are logged and replayed only when something needs that state. A first edit on a large workbook no longer waits seconds, and an editing session keeps its open-time memory. `Stylesheet` gains `same_state`.
