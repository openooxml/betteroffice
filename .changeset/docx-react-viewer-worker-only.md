---
"@betteroffice/docx-react": minor
---

Read-only and viewing editors keep documents only in the worker: `getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError` (use `readParagraphs` or `findText`), `getSelectionInfo` returns `null` (use `readSelectionInfo`), and unanswerable async calls, export included, reject. A worker that fails to open twice reports the new `DocxWorkerError` to `onError`.
