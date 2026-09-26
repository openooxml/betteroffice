---
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Resolve pictures stored only as SVG: a blip with no `r:embed` of its own takes its media from the Office SVG extension (`asvg:svgBlip`), for pictures and shape picture fills alike. A blip that carries both keeps its raster image.
