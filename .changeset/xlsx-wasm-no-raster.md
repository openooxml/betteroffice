---
"@betteroffice/xlsx": patch
---

`renderPng` without a canvas now rejects with `PngExportError` code `canvas-unavailable`; use `@betteroffice/xlsx-native` there. The vendored wasm module no longer declares `XlsxDocument.renderPng` or `renderRangePng`.
