# betteroffice-ooxml-raster

Shared tiny-skia helpers for the BetterOffice raster backends: PNG encoding,
hex and CSS color parsing, DrawingML geometry and glyph outline paths, RGBA
premultiplication and clip rectangles clamped to the target surface.

```rust
use ooxml_raster::{encode_png, parse_css_color};

let color = parse_css_color("rgba(0, 0, 0, 0.5)")?;
let png: Vec<u8> = encode_png(pixmap, width, height)?;
```

Used by [betteroffice-docx-raster](https://crates.io/crates/betteroffice-docx-raster),
[betteroffice-pptx-raster](https://crates.io/crates/betteroffice-pptx-raster),
[betteroffice-xlsx-raster](https://crates.io/crates/betteroffice-xlsx-raster)
and the VSDX raster backend. Each backend keeps its own color grammar and
error text on top of these helpers.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
