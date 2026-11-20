---
'@betteroffice/vsdx': minor
---

Breaking: `exportPng(pageIndex, scale)` is now `await renderPng({ page: pageIndex }, { scale })`, painted through the canvas. It rejects with `PngExportError` rather than substitute a font, or when no canvas exists. `paintPage` now draws letter spacing and underline.
