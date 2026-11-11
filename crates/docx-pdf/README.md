# betteroffice-docx-pdf

Native PDF export for BetterOffice DOCX display lists.

The writer consumes the renderer-agnostic `DisplayList` produced by
`docx-layout`; it never reads DOCX and does not reuse the VSDX PDF writer.

Supported today:

- page geometry, backgrounds, page borders (back/front z-order)
- positioned `GlyphRun` text with subset TrueType fonts, `ToUnicode`, widths
  and per-glyph placement
- fallback `Text` via Base-14 Helvetica (ASCII only)
- headers, footers and note regions as separate paint layers
- filled/stroked rectangles, lines, decorations and shape paths
- images as RGB XObjects with optional alpha soft masks, rotation and data URLs

Not yet covered: hyperlinks, tagged PDF/PDF-UA, gradients, SVG/EMF/JPEG XL
images, and crop rectangles that need source-pixel transformation.

```rust
use betteroffice_docx_pdf::export_pdf_with_fonts;
use docx_layout::display_list::DisplayList;
use ooxml_text::FontStore;

let bytes = export_pdf_with_fonts(&display_list, &font_store)?;
```

`export_pdf` is the no-font entry point for text/rect/line-only display lists.
`export_pdf_with_fonts_and_images` adds registered image bytes keyed by
relationship id.
