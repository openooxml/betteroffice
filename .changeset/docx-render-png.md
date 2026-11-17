---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
---

`YrsSession.renderPng({ page }, options?)` draws a page as PNG bytes with the editor's layout and painter, or natively with `@betteroffice/docx-native`. It rejects with `PngExportError` rather than paint a substitute font.
