---
'@betteroffice/docx': patch
'@betteroffice/fonts': patch
---

`renderPng` without `fontProvider` now loads the complete bundled fonts on the canvas and native engines. The native engine changes too: an installed Latin or international face that cannot load now rejects with `font-unavailable` instead of being skipped. Passing `fontProvider` keeps the old behaviour.
