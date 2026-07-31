---
"@betteroffice/rust-crates": patch
---

`betteroffice-docx` can now bound its own parsing and rasterize a page. `Document::open_with_limits` takes an explicit `ParseLimits`, so a host that ingests untrusted uploads caps paragraphs, tables, XML events and nesting itself instead of inheriting the parser's defaults; a document past any cap is refused rather than silently truncated.

Under the new opt-in `raster` feature, `Document::register_font` supplies font faces for shaping and `Document::render_png` rasterizes one display-list page to deterministic PNG bytes. Registration is budgeted the same way the PPTX facade budgets it — 256 faces, 32 MiB per font — and exceeding either budget is a typed `Error::ResourceLimit`, so a caller cannot quietly retain hundreds of megabytes of attacker-supplied fonts. Embedded images need no side table: the parser already resolves them to `data:` URLs against their own part's relationships, which the raster backend decodes directly, so a header image and a body image that share a relationship id stay distinct.

The feature is off by default because the raster backend is server-side only; the default build still compiles to `wasm32-unknown-unknown`. `Document` remains `Send + Sync` with fonts registered, so it stays usable from language bindings.
