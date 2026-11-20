---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
'@betteroffice/xlsx': minor
'@betteroffice/pptx': minor
---

DOCX worker recovery after a worker restart no longer fails, or brings back undone text, when undo or redo was used. Adds the `applyUpdate` command kind, which only recovery sends, and the required `applyVerbatimUpdate` engine method. XLSX and PPTX: a type addition only.
