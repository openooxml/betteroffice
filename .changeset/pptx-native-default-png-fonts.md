---
'@betteroffice/pptx-native': patch
---

`renderPng` works with no fonts passed: it uses the deck's embedded faces and the installed font packs, and honours `options.fonts`. `layoutSlide` results are unchanged by an export.
