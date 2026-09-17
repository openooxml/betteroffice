---
"@betteroffice/vsdx": minor
---

Export a diagram as SVG or PNG. `exportSvg()` returns one vector page per diagram page, sized from the PageSheet, with text kept as `<text>`; `exportPng(pageIndex, scale)` rasterizes one page at `scale` times the 96 dpi display list. Both walk the same ordered display list as the canvas and the PDF exporter, and both paint shape shadows.
