---
'@betteroffice/docx-react': patch
'@betteroffice/docx-i18n': patch
'@betteroffice/docx': minor
---

A cut (Ctrl+X or Shift+Delete) whose selection holds objects the clipboard cannot carry, such as notes, shapes, charts, equations, content controls or page and column breaks, is now refused with a notice and a recoverable `onError` report instead of deleting them. `DocxClipboardSlice` gains `leftOut`.
