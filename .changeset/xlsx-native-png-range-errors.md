---
'@betteroffice/xlsx-native': patch
'@betteroffice/xlsx': patch
---

`renderPng` with a bad range now rejects with the same `Error` and message on the native engine as on wasm (`bad range: ...`), instead of a plain string. Native range targets are read case as given, as on wasm, so lowercase A1 text is refused on both.
