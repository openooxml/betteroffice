---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

`renderPng` now paints with the faces passed in `fonts`, as the wasm engine does, and rejects with `font-unavailable` when a supplied face cannot be used. The option was previously ignored.
