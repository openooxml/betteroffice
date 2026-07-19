# betteroffice-pptx

The typed native Rust API for opening, inspecting, collaboratively editing,
laying out, and saving PPTX presentations. Parsed PresentationML and Yrs deck
state are exposed as Rust structs; native calls do not cross a JSON or
`JsValue` boundary.

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
# Ok::<(), betteroffice_pptx::Error>(())
```

The display list contains vector geometry, images, shaped text, caret data, and
hit-test metadata. Register at least one font face before rendering a slide that
contains text.

Saving currently byte-preserves the parsed source package. Yrs edits affect the
typed deck snapshot and rendered display lists, but are not yet projected back
into PresentationML parts. Persisting edited deck state is a lower-engine
follow-up. Added or removed slides and shapes therefore remain live editing and
collaboration operations until that projection exists.

The WASM crate surface remains available in `pptx-edit` for JavaScript clients;
this facade deliberately exposes the same native engine operations without its
JSON argument and result wrappers. The API is experimental and may change
before `0.1.0`.

## Support Matrix

| Capability | Native facade |
| --- | --- |
| Presentation, slide, master, layout, shape, text, theme, and media inspection | Yes |
| Yrs slide, shape, and text editing | Yes |
| Yrs v1 state vectors, diffs, updates, undo, and redo | Yes |
| Slide display lists and hit testing | Yes |
| Byte-preserving source package save | Yes |
| Persist Yrs edits into PresentationML | Follow-up |
