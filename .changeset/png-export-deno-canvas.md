---
'@betteroffice/docx': patch
---

DOCX PNG export rejects with `canvas-unavailable` in runtimes whose canvas has no 2D context, such as Deno, instead of failing while drawing, as XLSX and PPTX export already do.
