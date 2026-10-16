---
"@betteroffice/docx-react": minor
---

With `experimentalWorkerOpen`, read-only and viewing editors keep documents only in the worker: `getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError`, `getSelectionInfo` returns `null`, and unanswerable async calls reject. Opens and viewer renders failing on a replacement worker report the new `DocxWorkerError` to `onError` and show an alert.
