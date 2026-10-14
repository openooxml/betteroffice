# betteroffice-docx

Open, inspect, edit, lay out, and save DOCX documents from native Rust. Parsing
and serialization run on the OOXML model; paragraph edits run on the Yrs
editing core. Neither crosses a JSON or wasm boundary.

```rust
use betteroffice_docx::{Document, get_paragraph_text};

let mut document = Document::open(&docx_bytes)?;
let paragraph_id = document.paragraphs()[0].para_id.clone().unwrap();

document.replace_paragraph_text(&paragraph_id, "Updated in Rust")?;
assert_eq!(
    get_paragraph_text(document.paragraph(&paragraph_id).unwrap()),
    "Updated in Rust",
);

let saved = document.save()?;
```

Each paragraph ID addresses one body paragraph. A paragraph whose `w14:paraId`
repeats an earlier paragraph's carries a fresh ID that no part of the package
uses, which `save` writes once that paragraph is edited or the model is taken
through `model_mut`; until then it saves with its authored ID. A lookup that
still matches several paragraphs returns `None` or `Error::AmbiguousParagraph`.

`DocumentModel` exposes the body, sections, headers, footers, notes, styles,
numbering, relationships, media, and charts. `save` rewrites the parts the
engine owns and reuses the original package for the rest, so untouched parts
survive the round trip.

`open_with_limits` swaps the parser's default resource budget for a
caller-supplied `ParseLimits`, which is what a host ingesting untrusted uploads
wants. A document past any cap is refused, never truncated.

`export_structured` and `export_markdown` export the document, edits included,
as read-only structured content or Markdown with source anchors, and
`render_docx_markdown` renders exported content. `ExportOptions` requires a
`RevisionView` and selects stories (the body by default); omitted and
unsupported content is listed in the content's diagnostics, and options out of
range return `Error::Export`. `betteroffice-docx-edit`'s `EngineSession`
attaches a page map to a layout it computed itself.

`list_content_controls` and `find_content_controls` list the content controls of
the current model, edits included, with their tag, alias, type, lock, placement,
anchor and current text; ids and anchors address the returned snapshot. Filling
controls needs an `EditingDoc` session.

## Rendering

The opt-in `raster` feature adds native PNG rendering; the default build targets
`wasm32-unknown-unknown`.

```toml
betteroffice-docx = { version = "0.3", features = ["raster"] }
```

`render_png` takes a `DisplayList`, which `layout` returns alongside the typed
layout. `layout_input` is the measured projection the caller supplies.

```rust
use betteroffice_docx::{Document, ImageScope};

let mut document = Document::open(&docx_bytes)?;
document.register_font("Carlito", false, false, &font_bytes)?;
document.register_image(ImageScope::Body, "rId9", &logo_bytes)?;

let display_list = document.layout(layout_input)?.display_list;
let rendered = document.render_png(&display_list, 0)?;
// rendered.bytes, rendered.skipped_images
```

Faces are capped at 256 and fonts at 32 MiB each, images at 256 and 32 MiB
each, and a rendered page at 16384px per side and 16777216 pixels of area. Each
is a typed error rather than a truncated accept, and the page budget is checked
before any surface is allocated.

Images are budgeted by both the pixels and the bytes they decode to: 33554432
pixels for one image and 67108864 across one page, and 268435456 bytes for one
image and 536870912 across one page, charged from the declared extent and
colour depth before the decoder allocates. Pixels alone are not memory — a
16-bit source needs twice the buffer an 8-bit one of the same extent does. A
`data:` payload is bounded at 33554432 bytes and charged from its encoded
length before base64 expands it. A page decodes each resolved image once
however many primitives reference it, so its cost follows the budget rather
than the reference count.

A page also carries a work budget for what it allocates beyond its own
surface. Crop masks, clip surfaces and generated paths are charged in bytes —
32 per page pixel, and never under 32 MiB — because a crop mask and a clip
surface are both page-sized and a path grows with a number off the display
list. A wave, a dash pattern and a shape's geometry all expand into such a
path before anything clips them, so each is charged from the length it expands
over. Glyphs are counted instead of weighed, 1000000 across every run on the
page, and charged before the shaper runs: from the uppercased text where
`allCaps` is set, and from a tab leader's glyph times its repeat count.

A surface the page holds is charged its high-water mark, not once per use. One
clip surface and one crop-mask cache are alive at a time, so a page of distinct
cropped images costs the cache rather than a page-sized mask per reference. The
cache evicts the least recently used, because one cropped icon per table row
walks its geometries in a round robin. Exceeding the budget is an error: unlike
an image, the display list is the caller's own artifact and half a wave has no
partial form to fall back on.

`register_font` appends to the `family|bold|italic` fallback chain instead of
replacing it, so a second face for one family adds coverage for the glyphs the
first face lacks. `Presentation::register_font` replaces, deliberately: PPTX
resolves one face per family, DOCX resolves a chain.

## Images

Most embedded media arrives on the display list as a `data:` URL that
`docx-parse` already resolved against its owning part, and needs nothing
further. `register_image` supplies bytes keyed by owning part and relationship
ID. `render_png` reports unresolved references in `skipped_images`. Missing font
chains return an error.

## Editing and layout

- `replace_paragraph_text` rewrites a single-run paragraph, or returns
  `Error::UnsupportedParagraphEdit`; the re-exported `EditingDoc` exposes typed
  editing operations.
- Pagination accepts a measured `LayoutInput` and returns layout and a display
  list. `docx_edit::EngineSession` provides document lowering, measurement and
  retained layout.

Pre-1.0: the API may change between minor versions.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
