---
"@betteroffice/docx-react": patch
---

Without `experimentalWorkerOpen`, the editor keeps the drawn content of at most 192 pages near the viewport and rebuilds pages further away when they scroll back into view, so long documents use less memory. Printing still draws every page.
