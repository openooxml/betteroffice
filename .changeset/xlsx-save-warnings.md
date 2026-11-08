---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': patch
---

Adds `saveWithWarnings()` to XLSX workbook handles and sessions (`save_with_warnings()` in Python), returning the saved bytes with a `SaveWarning` for every part the save wrote differently from what the workbook holds: `{ code, part, element?, sheet?, message }`, the same shape on wasm, native, Python and Rust. Today the one code is `unwritable-reference`: the workbook holds a reference with a character XML cannot carry, so the save keeps that part as the file wrote it. No edit produces one today, and an ordinary save carries none. Plain `save()` returns the same bytes and also reports each warning, as a Node process warning with code `BETTEROFFICE_SAVE_WARNING` or a Python `RuntimeWarning`.
