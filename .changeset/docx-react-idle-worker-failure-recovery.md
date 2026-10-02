---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, an engine worker that fails after the document has loaded is replaced by a fresh one, so pages keep rendering; a repeated failure is reported through `onError`.
