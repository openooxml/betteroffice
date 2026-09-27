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

## SVG pictures

A picture whose bytes are SVG is rasterized natively with resvg/usvg instead of
the `image` crate. It draws shapes, paths, fills, strokes, linear and radial
gradients, clip paths, `<use>` and `<symbol>` instances, and stylesheets of
plain type, class and id rules. It does not draw text (`usvg`'s text feature is
off), embedded raster images (both href resolvers return `None`), markers,
filters, masks or patterns. A hyperlink or an embedded image leaves the rest of
the picture drawn; any other element outside that set declines the document.

The sandbox holds every decode to one envelope, `SVG_MEMORY_ENVELOPE` (100 MiB
beyond the output raster and its layers) and `SVG_TIME_ENVELOPE` (about a
second of one core), by construction: every bound below is a measured share of
it, and the shares are checked at compile time to sum within it.

- **Before `roxmltree`.** One pass over the bytes bounds nesting to
  `MAX_SVG_DEPTH`, attributes to `MAX_SVG_ELEMENT_ATTRIBUTES` per element,
  the `<` and `=` `roxmltree` reserves a node and an attribute for to
  `MAX_SVG_NODES` and `MAX_SVG_ATTRIBUTES`, and namespace declarations to
  `MAX_SVG_NAMESPACES`, and refuses any DTD.
- **Re-serialised.** `usvg` never reads the source. The document is written
  out again with only allowlisted elements and, on them, only allowlisted
  attributes in no namespace, each value re-emitted from the parse `usvg`
  makes of it: numbers, lengths, paints, transforms, path data and points in
  canonical form, `href` as it resolves, and stylesheets and `style`
  attributes as the declarations `simplecss` yields. Any attribute value but
  path data, points and `style` is held to `MAX_SVG_VALUE_BYTES`, and the
  sanitised document to `MAX_SVG_SANITIZED_BYTES`. Metadata, foreign markup,
  text, embedded images and anything in a gradient but its stops are left out,
  and a reference into them, to an id that cannot be written back unchanged,
  or to an existing element of another kind than the property converts is
  refused. So are context paint, a stop coloured `currentColor`, a gradient
  coordinate or radius in `em` or `ex`, and an attribute on the SVG, XLink or
  XML prefix other than `xlink:href`, `xlink:title`, `xml:space` and
  `xml:lang`. The sanitised bytes pass the same pre-parse scan as the source.
- **Before `usvg`.** The audit reads the sanitised document, exactly what
  `usvg` will. Every reference must name a fragment of the document, and the
  graph they form must be acyclic; a gradient chain links at most four. The
  tree every `<use>` and clip path expands into is sized first, to
  `MAX_SVG_EXPANDED_NODES` elements and `MAX_SVG_EXPANDED_BYTES` of markup,
  with `MAX_SVG_PATH_BYTES` per shape and `MAX_SVG_GRADIENT_STOPS` per
  gradient. Stylesheets are plain type, class and id rules within
  `MAX_SVG_STYLE_RULES`, `MAX_SVG_SELECTOR_PARTS` and
  `MAX_SVG_SELECTOR_BYTES`, and `MAX_SVG_STYLE_WORK` charges what `simplecss`
  spends on them, text rescans included; CSS may not say `inherit`. Gradient
  copies for the shapes that inherit one are charged to `MAX_SVG_PAINT_BYTES`,
  collecting them and the clip paths to `MAX_SVG_COLLECT_WORK`, and a dash
  list to every shape and `<use>` that may stroke with it. Every arc is bounded
  from its radii before `kurbo` subdivides it: at most 64 cubics, each weighed
  against `MAX_SVG_PATH_BYTES`, as are the arcs of circles, ellipses and
  rounded rectangles, whose radii must be absolute. Where anything is stroked,
  every shape is charged the pieces the stroker may emit as `usvg` strokes it
  to measure it, within `MAX_SVG_STROKE_VERBS`; curves and widths must stay
  within `MAX_SVG_STROKE_SPAN` tolerances, and rotations and skews are refused.
  The inherited-property lookups `usvg` makes through every ancestor, the
  values it parses anew, and the gradient chains and stops it converts again
  for every shape and `use` are charged to `MAX_SVG_INHERIT_WORK`.
- **Before `resvg`.** The converted tree is priced: group layers stack at most
  `MAX_SVG_LAYER_DEPTH` deep and are charged to the slide's image budget with
  the output raster, painted area stays within `MAX_SVG_OVERDRAW` times that
  raster, and painting work within `MAX_SVG_RENDER_WORK`: the edges the
  clipper cuts, the lines closing open contours and their sorting included,
  and each stroke's outline, which
  is stroked once here as `resvg` will stroke it. The raster renders at up to
  four times the document's intrinsic size and steps down when it would not
  fit, the slide's image budget left included; it and every layer stay within
  `MAX_SVG_RASTER_DIM` a side, one `tiny-skia` tile.
- **During both.** A panic in either library becomes a refusal.

A refused document is a skipped image like any other and carries nothing from
the document. A tiled SVG repeats at its intrinsic size, as a raster repeats at
its pixel size.

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
