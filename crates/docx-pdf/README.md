# betteroffice-docx-pdf

Native PDF export for BetterOffice DOCX display lists.

The writer consumes the renderer-agnostic `DisplayList` produced by
`docx-layout`; it never reads DOCX and does not reuse the VSDX PDF writer.

Supported today:

- page geometry, backgrounds, page borders (back/front z-order)
- positioned `GlyphRun` text with subset TrueType/CFF fonts, `ToUnicode`,
  widths and per-glyph placement
- per-character fallback across every registered face; Base-14 Helvetica only
  when the store has no usable outline face
- one CID per `(glyph, source text)` for TrueType fonts, with a generated
  `/CIDToGIDMap`, so an Arabic tooth/dot glyph shared by several letters still
  maps to the right character; non-base dot glyphs are marked with empty
  `/ActualText` and stay out of the text layer
- CFF fonts cannot duplicate CIDs: a glyph shared by several source characters
  keeps its first ToUnicode mapping (CJK subsets normally hit no duplicate)
- CFF2 faces are instanced and converted to TrueType outlines before embedding
- `TextRun` primitives carry `/ActualText`, so a scalar its chosen face lacks
  (for example Simplified Chinese in a Japanese chain) is still searchable when
  no covering face was registered
- headers, footers and note regions as separate paint layers
- filled/stroked rectangles, lines, decorations and shape paths
- images as RGB XObjects with optional alpha soft masks, crop, rotation,
  `flipH`/`flipV` and data URLs
- SVG pictures and EMF/EMF+/WMF metafiles rasterized through the shared
  bounded vector decoders (raster size and render work are capped); metafile
  text is laid out with the document's font store
- Flate-compressed page, image and font streams
- external link annotations from primitive `href`s, including rotated text,
  glyph runs, images and shapes; merged per target
- opt-in PDF/A-2u (ISO 19005-2 level U): embedded sRGB output intent, XMP
  `pdfaid` identification, matching document information, deterministic trailer
  IDs, nonzero glyph codes and Unicode mappings for every painted
  scalar. Standard exports remain plain PDF 1.7 and make no PDF/A claim.
- PDF/A-2u registers a bundled monochrome Noto Emoji fallback so emoji scalars
  are painted from a real outline instead of `.notdef`

Not yet covered: tagged PDF/PDF-UA, PDF/A-1, PDF/A-3, gradients, JPEG XL
images, named destinations for internal anchors (internal `#anchor` links are
skipped, not emitted as unresolved `/GoTo` actions), and color emoji
rendering.

```rust
use betteroffice_docx_pdf::export_pdf_with_fonts;
use docx_layout::display_list::DisplayList;
use ooxml_text::FontStore;

let bytes = export_pdf_with_fonts(&display_list, &font_store)?;
```

For archival output, use the facade's options API:

```rust
use betteroffice_docx::{Document, PdfConformance, PdfExportOptions, PdfMetadata};

let document = Document::open(&docx_bytes)?;
let pdf = document.export_pdf_with_options(&PdfExportOptions {
    conformance: PdfConformance::PdfA2u,
    metadata: PdfMetadata {
        title: Some("Contract".into()),
        ..PdfMetadata::default()
    },
})?;
```

PDF/A-2u requires a usable font store and refuses Base-14 or uncovered text
instead of writing a non-conforming file. A bitmap-only face is rejected rather
than synthesized into blank outlines, and bitmap faces are skipped when an
outline face can cover the same scalar. Invalid metadata dates are rejected;
absent fields are omitted from both `/Info` and XMP.

`export_pdf` is the no-font entry point for text/rect/line-only display lists.
`export_pdf_with_fonts_and_images` adds image bytes keyed by the scoped
relationship key described above.
