---
"@betteroffice/docx-react": patch
---

Stop eagerly re-parsing the document after opening; the model for `onChange`/`onContentChange`, save, export and `getDocument()` now builds lazily on first use.
