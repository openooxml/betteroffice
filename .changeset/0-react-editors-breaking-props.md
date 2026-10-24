---
"@betteroffice/docx-react": minor
"@betteroffice/xlsx-react": minor
"@betteroffice/pptx-react": minor
---

**Breaking:** `DocxEditorProps`, `XlsxEditorProps` and `PptxEditorProps` are unions, so interfaces can no longer extend them and `Omit` or `Pick` keep only shared props. Extend `DocxAsyncEditorProps`, `XlsxAsyncEditorProps` or `PptxAsyncEditorProps` instead. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
