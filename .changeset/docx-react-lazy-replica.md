---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, a read-only editor now loads its main-thread copy of the document only when something needs it, so a large document no longer freezes the page after it appears.
