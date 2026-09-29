---
"@betteroffice/docx-react": patch
---

Keep a canvas page's accessibility mirror and content-control overlay only while the page is in the page window (the viewport and two pages either side) or holds focus. Every page still keeps its sized canvas, so page geometry is unchanged. On a 267-page document the editor no longer builds and keeps about 630,000 mirror nodes for pages far from the viewport, or rebuilds them after every edit that moves their positions.
