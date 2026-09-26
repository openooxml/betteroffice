# betteroffice-pptx

Open, inspect, edit, collaborate on, and render PPTX presentations from native
Rust. Parsed PresentationML and Yrs deck state are plain Rust structs; calls do
not cross a JSON or `JsValue` boundary.

```rust
use betteroffice_pptx::{EditCtx, Presentation};

let mut presentation = Presentation::open(&pptx_bytes)?;
let snapshot = presentation.snapshot()?;
let slide = &snapshot.slides[0];
let shape = &slide.shapes[0];

presentation.move_shape(
    &EditCtx::local("example"),
    &slide.id,
    &shape.id,
    914_400,
    914_400,
)?;

presentation.register_font("Inter", false, false, &font_bytes)?;
let rendered = presentation.render_slide(0)?;
let saved = presentation.save()?;
```

Display lists carry vector geometry, images, shaped text, caret data, and
hit-test metadata. Register at least one font face before rendering a slide
that contains text.

`read_content`, `find_text`, `validate_edits` and `apply_edits` run
version-checked edit batches: every step resolves against the version the host
read, and the batch commits as one transaction or comes back as an
`EditRefusal` with nothing changed. The request types are the JSON contract of
the JavaScript core.

`export_structured` and `export_markdown` return the committed deck as
structured content or Markdown with the version it was read at;
`export_pptx_structured`, `export_pptx_markdown` and `render_pptx_markdown` do
the same for bytes and for content read earlier. Slides, shapes, stories,
paragraphs, tables and placeholders for unrepresented objects carry session or
snapshot anchors and source provenance, and every omission is a diagnostic.
Text anchors are `PptxAnchor::Range` batch targets.

`pptx-edit` keeps the wasm surface for JavaScript clients. This facade exposes
the same engine operations without its JSON argument and result wrappers.

`0.2.x`: the API may change before `1.0`.

## Limits

`save` writes Yrs edits back into PresentationML. Parts an edit did not touch
keep their exact source bytes; edited slides, notes, and comment parts are
patched at the XML level, so unmodeled markup — transitions, timing, hyperlinks,
fields, unknown attributes — survives; inserted and removed slides rewrite
`presentation.xml`, its relationships, and `[Content_Types].xml`. The ZIP
container is rebuilt, so the output is not byte-identical to the input even
without edits.

## Support matrix

| Capability | Native facade |
| --- | --- |
| Presentation, slide, master, layout, shape, text, theme, and media inspection | Yes |
| Yrs slide, shape, and text editing | Yes |
| Version-checked edit batches | Yes |
| Structured JSON and Markdown export with anchors | Yes |
| Yrs v1 state vectors, diffs, updates, undo, and redo | Yes |
| Slide display lists and hit testing | Yes |
| Part-preserving package save | Yes |
| Persist Yrs edits into PresentationML | Yes |

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
