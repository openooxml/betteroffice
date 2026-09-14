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

`pptx-edit` keeps the wasm surface for JavaScript clients. This facade exposes
the same engine operations without its JSON argument and result wrappers.

`0.0.x`: the API may change before `0.1.0`.

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
| Yrs v1 state vectors, diffs, updates, undo, and redo | Yes |
| Slide display lists and hit testing | Yes |
| Part-preserving package save | Yes |
| Persist Yrs edits into PresentationML | Yes |

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
