---
'@betteroffice/vsdx': patch
---

Vsdx text no longer collapses on Linux Chromium. The painter sets real-pixel font sizes with unrounded advances, so a registered face keeps its letters apart.
