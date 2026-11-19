---
'@betteroffice/docx-react': minor
---

`DocxEditor` gains `onReady(api)`: on the worker editor it is called once per document after its first page has painted, on a later turn, with the api the ref exposes, and never after the editor has unmounted or the file has changed.
