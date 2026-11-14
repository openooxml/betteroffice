---
'@betteroffice/rust-crates': minor
---

Enables PPTX raster support by default. For builds without it, set `default-features = false` and select font features explicitly; `wasm32` builds must do this, since the rasterizer and its embedded base Latin faces are native only.
