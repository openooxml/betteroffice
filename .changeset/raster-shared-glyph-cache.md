---
"@betteroffice/rust-crates": patch
---

The DOCX and PPTX raster crates share one glyph outline cache and glyph fill from betteroffice-ooxml-raster, and the VSDX raster reuses glyph outlines within a page. Rendered pixels are unchanged.
