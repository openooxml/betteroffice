---
'@betteroffice/rust-crates': patch
---

`betteroffice-docx-raster` turns on the bundled Latin faces through a default `base-latin` feature, so a crate that links the rasterizer can leave them out. Default behaviour is unchanged.
