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
let png = document.render_png(&display_list, 0)?;
```

Faces are capped at 256 and fonts at 32 MiB each, images at 256 and 32 MiB
each, a rendered page at 16384px per side and 16777216 pixels of area, and a
decoded image at 16384px per side and 67108864 pixels. Every budget is a typed
error rather than a truncated accept, and the render budgets are checked before
any surface is allocated.

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

An unresolved reference is skipped, matching the canvas backend's resolver —
one missing linked image must not blank the page around it. `register_image`
supplies the bytes where skipping is not what you want. It is keyed by owning
part, because a header and the body can both use `rId9` for different media.

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
