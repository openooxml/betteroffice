# betteroffice-docx-pdf

Native PDF export for BetterOffice DOCX display lists.

The writer consumes the renderer-agnostic `DisplayList` produced by
`docx-layout`; it never reads DOCX and does not reuse the VSDX PDF writer.

Supported today:

- page geometry, backgrounds, page borders (back/front z-order)
- positioned `GlyphRun` text with subset TrueType/CFF fonts, `ToUnicode`,
  widths and per-glyph placement
- per-character fallback across every registered face; Base-14 Helvetica only
  when the store has no fonts at all
- headers, footers and note regions as separate paint layers
- filled/stroked rectangles, lines, decorations and shape paths
- images as RGB XObjects with optional alpha soft masks, crop, rotation and data URLs
- Flate-compressed page, image and font streams
- external link annotations from primitive `href`s, merged per target

Not yet covered: tagged PDF/PDF-UA, gradients, SVG/EMF/JPEG XL images,
named destinations for internal anchors (internal `#anchor` links are
skipped, not emitted as unresolved `/GoTo` actions), and color emoji rendering (missing emoji glyphs
still produce a searchable .notdef with a coverage warning).

```rust
use betteroffice_docx_pdf::export_pdf_with_fonts;
use docx_layout::display_list::DisplayList;
use ooxml_text::FontStore;

let bytes = export_pdf_with_fonts(&display_list, &font_store)?;
```

`export_pdf` is the no-font entry point for text/rect/line-only display lists.
`export_pdf_with_fonts_and_images` adds image bytes keyed by the scoped
relationship key described above.
