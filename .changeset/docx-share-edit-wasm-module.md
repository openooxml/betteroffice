---
"@betteroffice/docx": patch
---

`preloadDocxEngine()` and `experimentalPrewarm` now download and compile the editing engine once and share it with the preloaded worker, with or without `experimentalWorkerOpen`.
