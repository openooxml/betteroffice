---
'@betteroffice/docx': patch
'@betteroffice/fonts': minor
---

Without `fontProvider` or configured fonts, PNG export uses the complete bundled fonts on both engines, native included: an installed Latin or international face that cannot load rejects with `font-unavailable`. Aborting while fonts load rejects at once. `fontProvider` keeps the old behaviour. Fonts adds `isMissingFontAddon`.
