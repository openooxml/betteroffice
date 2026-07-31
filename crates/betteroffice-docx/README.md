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

```rust
let mut document = Document::open(&docx_bytes)?;
document.register_font("Carlito", false, false, &font_bytes)?;
let png = document.render_png(&display_list, 0)?;
```

Faces are capped at 256 and fonts at 32 MiB each; either budget is a typed
error. Embedded images travel on the display list as `data:` URLs already
resolved against their owning part, so rendering needs no image side table.

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
