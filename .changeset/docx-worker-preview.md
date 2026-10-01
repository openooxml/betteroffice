---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen` and `previewFirstPage`, the first-page preview now opens and lays out in the resident worker, so the main thread parses and lays out nothing before the first page shows.
