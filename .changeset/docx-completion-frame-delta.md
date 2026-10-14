---
"@betteroffice/docx-react": patch
---

When a long document finishes laying out after its first pages appear, only the pages that changed are sent to the main thread instead of every page. Applies with and without `experimentalWorkerOpen`.
