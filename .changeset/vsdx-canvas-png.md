---
'@betteroffice/vsdx': minor
---

Breaking: `exportPng(pageIndex, scale)` is now `await renderPng({ page: pageIndex }, { scale })`, painted through the canvas. It rejects with `PngExportError` rather than substitute a font, and without a canvas names `@betteroffice/vsdx-native`. `paintPage` now draws letter spacing and underline.
