---
"@betteroffice/pptx-react": patch
---

`save()` on the default editor's `onReady` API now returns a promise with the presentation bytes, like XLSX and DOCX. `saveAsync()` keeps working as a deprecated alias.
