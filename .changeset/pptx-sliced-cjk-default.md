---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
---

PowerPoint editors load Chinese, Japanese, and Korean fonts in slices by default in Chromium and Firefox with worker rendering, and use whole fonts in other browsers. If slices fail, browser fallback keeps text visible.
