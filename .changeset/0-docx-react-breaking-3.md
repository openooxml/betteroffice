---
"@betteroffice/docx-react": minor
---

**Breaking:** On the default editor, synchronous `getDocument`, `getPageContent` and `findInDocument` throw (`DocxAsyncOnlyError` in read-only and viewing sessions). Use the async `readParagraphs` or `exportStructuredWithPages` for content and `findParagraphs` for search.
