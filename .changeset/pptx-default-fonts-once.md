---
'@betteroffice/pptx': patch
---

The PowerPoint editor no longer sends a second copy of every default font's bytes to the main thread with its updates; only the browser font install receives them, so the page holds each default face once.
