---
'@betteroffice/docx-react': minor
---

`DocxEditor` gains `onReady(api)`, as the xlsx and pptx editors have: on the worker editor it is called once per document, with the api the ref exposes, when the opened document first paints and that api is usable, not at a `previewFirstPage` preview.
