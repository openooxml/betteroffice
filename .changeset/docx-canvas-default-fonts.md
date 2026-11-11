---
'@betteroffice/docx': patch
'@betteroffice/fonts': minor
---

Without `fontProvider`, PNG export uses the complete bundled fonts on both engines. On the native engine, an unloadable installed Latin or international face rejects with `font-unavailable` instead of being skipped. Aborted exports reject at once, not after font loads. Passing `fontProvider` keeps the old behaviour.
