---
"@betteroffice/docx-react": patch
---

Stop eagerly re-parsing the document after opening. The model for `onChange`/`onContentChange` now builds once the first pages appear, only if such a listener exists; save, export and `getDocument()` still build it on first use.
