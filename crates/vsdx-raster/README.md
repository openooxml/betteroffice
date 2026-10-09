# betteroffice-vsdx-raster

Server-side raster backend: paints a VSDX display-list page to PNG through
tiny-skia. Twin of the browser canvas painter in
`packages/vsdx/src/render/canvas.ts`, which stays the on-screen reference.

Text is set in the face registered for each run's family, else a registered
Arial or `sans-serif` face, else the vendored Carlito Regular shared with
`crates/xlsx-raster/assets/` (OFL, see `THIRD-PARTY-NOTICES.md`). No system
fonts are read, so output is identical on every machine. Bold and italic are
synthesized. Gradients render as linear blends, and images decode from JPEG and
PNG.
