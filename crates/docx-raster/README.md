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

## Budgets

The entry points bound themselves, so a consumer of this crate is bounded
without a facade in front of it. A page past `MAX_PAGE_DIM` per side or
`MAX_PAGE_PIXELS` of area is refused before any surface is allocated. Beyond
that surface, one page render may spend `MAX_PAGE_SCRATCH_BYTES_PER_PIXEL` per
page pixel (never under `MIN_PAGE_SCRATCH_BYTES`) on crop masks, clip surfaces
and generated paths, and may paint `MAX_PAGE_GLYPHS` glyphs across every run on
it. A wave, a dash pattern and a shape's geometry are all charged from the
length they expand a path over, before the expansion allocates. Images are
charged both `MAX_IMAGE_PIXELS`/`MAX_PAGE_IMAGE_PIXELS` and
`MAX_IMAGE_BYTES`/`MAX_PAGE_IMAGE_BYTES`, and a `data:` payload is bounded at
`MAX_DATA_URL_BYTES` before base64 expands it.

A surface the page holds is charged its high-water mark rather than once per
use, since one clip surface and one crop-mask cache are alive at a time. The
crop-mask cache evicts the least recently used, so repeating crop geometries
down a page reuse their filled masks.

An image over budget is skipped and counted in `RenderedPage::skipped_images`;
everything else over budget is an error.

## Image keys

`ImageMap` is keyed by `scoped_image_key`, so a header image and a body image
that share a relationship id stay distinct. A body image also resolves under a
bare relationship id, which is how this crate first shipped; a header does not,
because reaching another part's id is what scoping exists to prevent. Nor does
a body relationship id that spells another part's scoped key.

`render_png` returns bytes alone and cannot report a skipped image. Use
`render_page` where that matters.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
