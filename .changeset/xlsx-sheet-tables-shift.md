---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Tables stay on their own sheets when a sheet is inserted or deleted, so structured references keep reading the right cells; deleting a sheet drops its tables and turns references to them into `#REF!`. Undo restores them via a new internal `setTables` op.
