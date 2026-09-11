# betteroffice-pptx-raster

The tiny-skia backend that paints a PPTX slide display list to PNG. Server-side
twin of the browser's canvas replayer, and the CPU reference the native viewer
diffs its GPU output against.

```rust
use pptx_raster::{AssetMap, RenderOptions, RenderResources, render_slide};

let images: AssetMap<'_> = presentation
    .media()
    .iter()
    .map(|part| (part.part_path.as_str(), part.bytes.as_slice()))
    .collect();
let resources = RenderResources::new(renderer.fonts(), &images);
let png = render_slide(&display_list, &resources, &RenderOptions::default())?;
```

Most callers want the facade instead — `betteroffice-pptx` with the `raster`
feature gives you `Presentation::render_png`, which resolves media out of the
package for you.

PNG encoding uses fixed settings, so identical inputs produce byte-identical
output. `tests/golden.rs` byte-compares every scenario against a committed PNG;
regenerate deliberately with:

```bash
GOLDEN_UPDATE=1 cargo test -p betteroffice-pptx-raster
```

## Fonts

Nothing is embedded. Text is painted from the `PositionedGlyph` runs the layout
pass already placed, so this crate shapes nothing — it resolves each run's
`font_id` against the `FontStore` you hand it and fills the outline. Register
faces on the `SlideRenderer` (or the `Presentation`) before laying the slide out.

## Never in the wasm build

Decoding pictures needs the `image` crate, so `src/lib.rs` refuses to compile for
`wasm32`. The browser gets its PNG from `slideToPng` in `@betteroffice/pptx`,
which drives the canvas replayer and `canvas.toBlob()` instead.

## What the display list does not carry

These are gaps upstream of this crate, in the contract `pptx-render` emits, so
the PNG can only be as faithful as what the canvas backend already draws:

- **Picture crops.** `PictureCrop` (`srcRect`) is parsed but dropped by the
  layout pass, so a cropped picture paints stretched to its frame.
- **Tables.** They arrive as dashed `Placeholder` boxes labelled `"Table"`; the
  parsed cell content is never laid out.
- **Effects and alpha.** There is no shadow, glow, reflection, soft edge, or
  opacity in the contract, and colors resolve to `#rrggbb` with no alpha.
- **Pattern and picture fills.** `Paint` is only `Solid` or `Gradient`.
- **Dash patterns.** A stroke's dash collapses to one boolean, so the specific
  OOXML pattern is lost; this crate synthesizes the same dashes the canvas
  backend does, keeping the two in agreement.

Unlike `docx-raster`, which errors on anything it cannot reproduce faithfully,
these are absences in the input rather than fields being refused, so this crate
paints what is there. Only images degrade at render time: one that is missing,
undecodable, or over budget is skipped and counted in
`RenderedSlide::skipped_images`.
