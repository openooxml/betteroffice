---
"@betteroffice/docx": patch
---

`preloadDocxEngine()` and `experimentalPrewarm` compile the engine once and share it with the preloaded worker, which retries an interrupted download.
