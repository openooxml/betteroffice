---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

With `experimentalWorkerOpen`, rendering on the main thread after a worker fallback prepares only pages near the viewport, so large documents keep working while scrolling.
