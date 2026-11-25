---
'@betteroffice/pptx': patch
'@betteroffice/fonts': patch
---

The presentation editor now loads fallback fonts for scripts like Arabic, Hebrew, Thai and Devanagari by default and skips repeated downloads for unavailable packs. Simplified Arabic and Traditional Arabic request the same fallback face as PNG export.
