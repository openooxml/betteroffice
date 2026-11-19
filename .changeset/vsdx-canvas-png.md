---
'@betteroffice/vsdx': minor
---

Breaking: `exportPng(pageIndex, scale)` is now `await renderPng({ page: pageIndex }, { scale })`, painted through the canvas. It rejects with `PngExportError` rather than substitute a font, and without OffscreenCanvas or a DOM canvas says to run it in a browser or a web worker. `paintPage` now draws letter spacing and underline.
