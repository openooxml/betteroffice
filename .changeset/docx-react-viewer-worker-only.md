---
"@betteroffice/docx-react": minor
---

Read-only and viewing editors keep the document only in the document worker, with no main-thread copy: their `getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError` (use `readParagraphs`, `exportStructuredWithPages` and `findParagraphs` or `findText`), `getSelectionInfo` returns `null` (use `readSelectionInfo`), and async ref calls the worker cannot answer, export included, reject instead of falling back to the main thread. A document worker that fails to open or render after one restart now reports the new `DocxWorkerError` to `onError` in every session, without a main-thread fallback.
