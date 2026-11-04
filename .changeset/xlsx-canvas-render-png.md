---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/agents': patch
---

Breaking: workbook `renderPng(viewport)` and `renderRangePng(options)` become `await renderPng({ sheet, viewport } | { sheet, range }, { scale, maxWidth, maxHeight })`, painted with the editor's painter on a canvas. It rejects with `PngExportError` rather than substitute a font.
