---
'@betteroffice/xlsx-native': patch
'@betteroffice/xlsx': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

`renderPng` with a bad range now rejects with the same `Error` and `bad range: ...` message on native as on wasm, not a plain string. A1 ranges are read in either case on every engine, as Excel does: `a1:b2` is `A1:B2`.
