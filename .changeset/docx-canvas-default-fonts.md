---
'@betteroffice/docx': patch
'@betteroffice/fonts': minor
---

Without `fontProvider`, PNG export uses the complete bundled fonts on both engines, native included: an installed Latin or international face that cannot load rejects with `font-unavailable` instead of being skipped. An aborted export rejects at once. Passing `fontProvider` keeps the old behaviour.
