---
'@betteroffice/docx-react': minor
---

`DocxEditor` gains `onReady(api)`, as the xlsx and pptx editors have: on the worker editor it is called once per document, right after `onFirstPagePainted`, with the api the ref exposes.
