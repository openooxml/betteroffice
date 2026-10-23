---
"@betteroffice/docx-react": minor
---

**Breaking:** On the default editor, every synchronous ref member the worker-only editor does not implement, such as `getDocument`, `getSelectionInfo`, `addComment` and `applyFormatting`, throws `DocxSyncApiError` (`getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError` in read-only and viewing sessions). Replacements: https://docs.betteroffice.dev/docs/migrating-to-0-5.
