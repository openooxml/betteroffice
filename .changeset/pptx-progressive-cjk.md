---
'@betteroffice/fonts': patch
'@betteroffice/pptx': patch
'@betteroffice/pptx-react': patch
---

New `experimentalSlicedFonts` editor option, off by default: in Chromium and Firefox it loads only the Chinese, Japanese or Korean outlines a slide needs. In `@betteroffice/fonts`, `ensureText` and `prefetch` also load the characters shaping composes text into.
