---
"@betteroffice/docx-react": patch
---

Exports `DocxEditorNotReadyError` (`code: "editor-not-ready"`), which ref reads, `applyEdits` and `save()` reject with before a document has finished opening, so hosts can check for it with `instanceof`.
