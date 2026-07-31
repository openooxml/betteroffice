---
"@betteroffice/rust-crates": patch
---

`betteroffice-docx` can now bound its own parsing and rasterize a page. `Document::open_with_limits` takes an explicit `ParseLimits`, so a host that ingests untrusted uploads caps paragraphs, tables, XML events and nesting itself instead of inheriting the parser's defaults; a document past any cap is refused rather than silently truncated.

Under the new opt-in `raster` feature, `Document::register_font` supplies font faces for shaping and `Document::render_png` rasterizes one display-list page to deterministic PNG bytes. Registration is budgeted the same way the PPTX facade budgets it — 256 faces at 32 MiB each, plus 256 images at 32 MiB each, so 16 GiB retainable in the worst case — and exceeding any budget is a typed `Error::ResourceLimit`.

The display list is budgeted too, since it is the one input this API takes from outside. A page over 16384px per side or 16777216 pixels of area is refused before any surface is allocated, mirroring the XLSX facade. Images are budgeted by the pixels they decode to rather than by the bytes they arrive as: one image may decode up to 33554432 pixels and one page 67108864 across every image on it, charged from the declared extent before the decoder allocates. A page decodes each resolved image once however many primitives reference it, so a page's decode cost is bounded by the budget rather than by the number of references.

Most embedded media needs no side table: the parser resolves it to a `data:` URL against its own part's relationships, which the raster backend decodes directly. Media the parser could not resolve — an external or dangling relationship, an image outside `word/media/`, a picture watermark's bare `rId` — reaches the backend unresolved. `Document::register_image` supplies the bytes for those, keyed by owning part so a header image and a body image that share a relationship id stay distinct.

An image the backend will not draw is skipped rather than failing the page, matching the canvas backend's resolver: an unresolved reference, bytes that will not decode, and an image past a pixel budget all leave a hole instead of taking the page down. `render_png` returns that skipped count alongside the bytes, so a caller can log or reject on it.

The feature is off by default because the raster backend is server-side only; the default build still compiles to `wasm32-unknown-unknown`. `Document` remains `Send + Sync` with fonts registered, so it stays usable from language bindings.
