# betteroffice-docx

The typed native Rust API for opening, inspecting, editing, laying out, and
saving DOCX documents. Parsing and serialization use the native OOXML model,
while paragraph edits use the Yrs-backed editing core without crossing a JSON
or wasm boundary.

```rust
use betteroffice_docx::{Document, get_paragraph_text};

let mut document = Document::open(&docx_bytes)?;
let paragraph_id = document.paragraphs()[0].para_id.clone().unwrap();

document.replace_paragraph_text(&paragraph_id, "Updated in Rust")?;
assert_eq!(get_paragraph_text(document.paragraph(&paragraph_id).unwrap()), "Updated in Rust");

let saved = document.save()?;
# Ok::<(), betteroffice_docx::Error>(())
```

`DocumentModel` exposes the body, sections, headers, footers, notes, styles,
numbering, relationships, media, and charts. Saving replaces engine-owned DOCX
parts while retaining the original package as the source for unowned parts.

The high-level paragraph replacement currently accepts plain single-run
paragraphs. The complete native `EditingDoc` and its typed operation vocabulary
are re-exported for advanced editing workflows.

Pagination accepts a typed, already-measured `LayoutInput` and returns both the
typed layout and body display list. Native DOCX-model lowering and measurement
are not yet exposed by the lower crates, so callers must currently provide that
projection.

Publishing remains disabled until `docx-parse`, `docx-layout`, and `docx-edit`
are publishable dependencies. The API is experimental and may change before
`0.1.0`.
