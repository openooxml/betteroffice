---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

Build each canvas page's full accessibility mirror only near the viewport, keeping just the links, notes and header cells of pages farther away. Tab, note links and `RenderedDomContext.findElementsForRange()` still reach every page.
