---
'@betteroffice/pptx': patch
'@betteroffice/fonts': patch
---

The presentation editor loads fallback fonts for scripts like Arabic, Hebrew, Thai and Devanagari, using the same faces as PNG export. When a font pack is unavailable, it warns once and skips repeated downloads.
