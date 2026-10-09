---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

In viewer sessions `listContentControls`, `findContentControls` and the built-in Find read from the document worker without a main-thread document copy, and tracked-change accept and reject commands refuse at once.
