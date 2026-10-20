---
"@betteroffice/pptx-react": minor
---

**Breaking:** `onReady` receives `PptxAsyncEditorApi`: use `handleAsync()` instead of `handle` and `saveAsync()` instead of `save()`, and await `goToSlide`, `selectText`, `clearSelection`, `getPositionAtPoint`, `refresh` and `refreshProposals`, which now return promises.
