---
'@betteroffice/vsdx': minor
---

Breaking: `exportPng(pageIndex, scale)` is now `await renderPng({ page: pageIndex }, { scale })`. It paints through the canvas and rejects with `PngExportError` rather than substitute a font that cannot load. Without a canvas it rejects and names `@betteroffice/vsdx-native`.
