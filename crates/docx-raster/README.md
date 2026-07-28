# betteroffice-docx-raster

The native raster backend: paints one
[betteroffice-docx-layout](https://crates.io/crates/betteroffice-docx-layout)
display-list page to PNG through tiny-skia. It uses the same font store and
fallback chains as layout, resolves embedded image relationship IDs from
caller-provided bytes, and never enters the wasm build.

```rust
use docx_raster::{RenderResources, render_png};

let resources = RenderResources::new(&fonts, &font_chains, &images);
let png: Vec<u8> = render_png(&display_list, 0, &resources)?;
```

The renderer refuses missing resources and visual effects it cannot reproduce
faithfully. PNG encoding uses fixed settings, so identical inputs produce
byte-identical output.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
