---
'@betteroffice/fonts': minor
'@betteroffice/docx-react': minor
'@betteroffice/docx': patch
---

Adds `preloadFontScripts` to the worker editor: the bundled sliced fallback face for those scripts loads whole in the background, so typed text does not wait for its outlines. Off by default; progressive faces also gain `chunks()`, `ensureChunks()` and `chunksOfGlyphs()`.
