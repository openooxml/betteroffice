---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

Keep a canvas page's accessibility mirror and content-control overlay only while the page is in the page window (the viewport and two pages either side), holds focus, or is needed elsewhere. Tab, a link to a note, and `RenderedDomContext.findElementsForRange()` build a far page's chrome on demand. A page in the window rebuilds its chrome at once after an edit, and a content-control overlay always does. Every page still keeps its sized canvas, so page geometry is unchanged. On a 267-page document the editor no longer builds and keeps about 630,000 mirror nodes for pages far from the viewport, or rebuilds them after every edit that moves their positions.
