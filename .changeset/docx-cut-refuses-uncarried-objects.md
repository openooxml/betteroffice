---
'@betteroffice/docx-react': patch
'@betteroffice/docx-i18n': patch
'@betteroffice/docx': minor
---

A cut whose selection holds objects the clipboard cannot carry, such as notes, shapes, charts, equations or content controls, is now refused with a notice and a recoverable `onError` report instead of deleting them. `DocxClipboardSlice` gains `leftOut`, naming those kinds.
