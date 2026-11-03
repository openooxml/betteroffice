---
'@betteroffice/rust-crates': minor
---

Enables PPTX raster support by default. Set `default-features = false` and select desired font features explicitly to keep builds without raster support. Builds for `wasm32` must set `default-features = false` and enable the font features they need, since the rasterizer and its embedded base Latin faces are native only.
