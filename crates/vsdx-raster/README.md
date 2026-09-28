# betteroffice-vsdx-raster

Server-side raster backend: paints a VSDX display-list page to PNG through
tiny-skia. Twin of the browser canvas painter in
`packages/vsdx/src/render/canvas.ts`, which stays the on-screen reference.

Text is always set in the vendored Carlito Regular shared with
`crates/xlsx-raster/assets/` (OFL, see `THIRD-PARTY-NOTICES.md`), so output is
identical on every machine with no system font access. Bold and italic are
synthesized. Gradients render as linear blends, and images decode from JPEG and
PNG.
