---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Render a slide at the page extent PowerPoint exports. A slide's page box is a whole number of points, so `p:sldSz` snaps there before the pixel scale, and the canvas backing store now covers a fractional extent by rounding up, matching what the raster path already does. An A4 deck whose `p:sldSz` is 10691813×7559675 EMU renders 1755×1240 px at 150 DPI, the size of PowerPoint's own PDF page, instead of coming up a pixel short.
