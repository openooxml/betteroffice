---
"@betteroffice/pptx": minor
"@betteroffice/pptx-react": patch
---

PowerPoint editors download Chinese, Japanese, and Korean fonts in slices by default in Chromium and Firefox with worker rendering, and use browser fallback if slices fail. Other browsers use whole faces, and setting `experimentalSlicedFonts` to `false` switches to whole faces.
