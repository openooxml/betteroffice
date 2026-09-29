---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

Large documents now keep each page's accessible text in the DOM only near the viewport, on by default. Tab, links, content controls and `RenderedDomContext.findElementsForRange()` still reach every page.
