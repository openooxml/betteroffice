---
"@betteroffice/rust-crates": patch
---

The DOCX, PPTX, XLSX and VSDX raster crates share their PNG encoding, color parsing, path building and glyph outline helpers through the new betteroffice-ooxml-raster crate. Rendered pixels are unchanged.
