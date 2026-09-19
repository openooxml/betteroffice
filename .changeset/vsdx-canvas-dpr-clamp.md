---
'@betteroffice/vsdx': patch
'@betteroffice/vsdx-react': patch
---

Clamp the canvas backing store to the browser's maximum side length and to a per-canvas area budget, so a large page at high zoom on a high-density display no longer requests a canvas past that limit or allocates several hundred MiB for one page.
