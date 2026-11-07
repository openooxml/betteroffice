---
'@betteroffice/xlsx': patch
---

`renderPng` in the browser without `fonts` now paints with the bundled faces the native engine draws the sheet's fonts with, instead of the browser's fonts, so canvas and native PNGs match. A face that cannot load rejects with `PngExportError` (`font-unavailable`).
