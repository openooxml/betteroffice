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

## Limits

- `replace_paragraph_text` takes single-run paragraphs. Richer editing goes
  through the re-exported `EditingDoc` and its typed operation vocabulary.
- Pagination takes an already-measured `LayoutInput` and returns the typed
  layout plus the body display list. The lower crates do not yet expose
  DOCX-model lowering and measurement, so callers supply that projection.

`0.0.x`: the API may change before `0.1.0`.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
