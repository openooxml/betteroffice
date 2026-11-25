---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
---

PowerPoint editors in Chromium and Firefox load Chinese, Japanese, and Korean fonts in slices by default, so a slide downloads only the glyphs it shows. Other browsers keep whole faces. Set `experimentalSlicedFonts` to `false` to always load whole faces.
