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

`DocumentModel` exposes the body, sections, headers, footers, notes, styles,
numbering, relationships, media, and charts. `save` rewrites the parts the
engine owns and reuses the original package for the rest, so untouched parts
survive the round trip.

`open_with_limits` swaps the parser's default resource budget for a
caller-supplied `ParseLimits`, which is what a host ingesting untrusted uploads
wants. A document past any cap is refused, never truncated.

## Rendering

The `raster` feature is opt-in; the raster backend is server-side only and the
default build still targets `wasm32-unknown-unknown`.

```toml
betteroffice-docx = { version = "0.0.4", features = ["raster"] }
```

`render_png` takes a `DisplayList`, which `layout` returns alongside the typed
layout. `layout_input` is the measured projection the caller supplies; see
Limits below for why.

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
further. Media the parser could not resolve does not: an external (`r:link`)
or dangling relationship, an image outside `word/media/`, and a picture
watermark's bare `rId` all reach the backend unresolved.

An image the backend will not draw is skipped, matching the canvas backend's
resolver — one missing linked image, or one past a budget, must not blank
the page around it. `render_png` reports how many references it skipped, so a
caller that wants a whole page can reject on it. `register_image` supplies the
bytes where skipping is not what you want. It is keyed by owning part, because
a header and the body can both use `rId9` for different media.

Images skip and fonts do not, deliberately. An image is one element of a page
and its absence is a bounded hole a caller can see in `skipped_images` and act
on. A missing font chain is not bounded: every run in that family disappears,
and a page of invisible text reports the same success as a page of text. The
asymmetry is which failures leave a signal a caller can act on.

## Limits

- `replace_paragraph_text` takes single-run paragraphs. Richer editing goes
  through the re-exported `EditingDoc` and its typed operation vocabulary.
- Pagination takes an already-measured `LayoutInput` and returns the typed
  layout plus the body display list. The lower crates do not yet expose
  DOCX-model lowering and measurement, so callers supply that projection.
- `render_png` takes a `DisplayList` for that same reason: it is the last
  artifact on the pipeline the Rust crates can produce end to end. `DisplayList`
  deserializes, so a binding hands over the JSON its layout pass already emits.

`0.0.x`: the API may change before `0.1.0`.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
