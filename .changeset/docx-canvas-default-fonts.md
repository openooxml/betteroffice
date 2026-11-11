---
'@betteroffice/docx': patch
'@betteroffice/fonts': patch
---

`renderPng` without `fontProvider` now loads the complete bundled fonts on the canvas and native engines. The native engine changes too: a face that cannot load now rejects with `font-unavailable` on both engines instead of being skipped. Passing `fontProvider` keeps the old behaviour.
