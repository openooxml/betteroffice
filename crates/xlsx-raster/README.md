# betteroffice-xlsx-raster

The native raster backend: paints a
[betteroffice-xlsx-render](https://crates.io/crates/betteroffice-xlsx-render)
display list to PNG through tiny-skia. Server-side twin of the browser's canvas
backend, and never part of the wasm build.

```rust
use xlsx_raster::render_png;

let png: Vec<u8> = render_png(&display_list)?;
```

Because both backends consume the same display list, a sheet rendered on the
server matches what the browser paints.

Reachable from [betteroffice-xlsx](https://crates.io/crates/betteroffice-xlsx)
through its default `raster` feature.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
