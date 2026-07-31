---
"@betteroffice/rust-crates": patch
---

`betteroffice-docx` can now bound its own parsing and rasterize a page. `Document::open_with_limits` takes an explicit `ParseLimits`, so a host that ingests untrusted uploads caps paragraphs, tables, XML events and nesting itself instead of inheriting the parser's defaults; a document past any cap is refused rather than silently truncated.

Under the new opt-in `raster` feature, `Document::register_font` supplies font faces for shaping and `Document::render_png` rasterizes one display-list page to deterministic PNG bytes. Registration is budgeted the same way the PPTX facade budgets it — 256 faces, 32 MiB per font, so 8 GiB retainable in the worst case — and exceeding either budget is a typed `Error::ResourceLimit`. The display list is budgeted too, since it is the one input this API takes from outside: a page over 16384px per side or 16777216 pixels of area is refused before any surface is allocated, mirroring the XLSX facade, and a decoded image over 16384px per side or 67108864 pixels is refused before the decoder allocates its buffer.

Most embedded media needs no side table: the parser resolves it to a `data:` URL against its own part's relationships, which the raster backend decodes directly. Media the parser could not resolve — an external or dangling relationship, an image outside `word/media/`, a picture watermark's bare `rId` — reaches the backend unresolved, and is now skipped rather than failing the whole page. `Document::register_image` supplies the bytes for those, keyed by owning part so a header image and a body image that share a relationship id stay distinct.

The feature is off by default because the raster backend is server-side only; the default build still compiles to `wasm32-unknown-unknown`. `Document` remains `Send + Sync` with fonts registered, so it stays usable from language bindings.
