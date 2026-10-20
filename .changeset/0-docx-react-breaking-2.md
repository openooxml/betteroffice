---
"@betteroffice/docx-react": minor
---

**Breaking:** On the default editor, passing `onChange(Document)` reports an error through `onError` and renders that error instead of the editor. Use `onDocumentChange({ version })` with async reads such as `readParagraphs`, and the ref's `save()`.
