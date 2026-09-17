# betteroffice-vsdx-raster

Server-side raster backend: paints a VSDX display-list page to PNG through
tiny-skia. Twin of the browser canvas painter in
`packages/vsdx/src/render/canvas.ts`, which stays the on-screen reference.

Text is always set in the vendored Carlito Regular shared with
`crates/xlsx-raster/assets/` (OFL, see `THIRD-PARTY-NOTICES.md`), so output is
identical on every machine with no system font access. Run families, small
caps, superscripts and subscripts do not survive; bold and italic are
synthesized. Gradients render as linear blends; images decode from JPEG and
PNG only — anything else keeps a dashed placeholder and is counted in
`skipped_images`.
