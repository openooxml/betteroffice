---
'@betteroffice/rust-crates': minor
---

Consolidates viewport PNG exports into `render_png(PngTarget, &RenderOptions)`. Replace `render_png(&viewport)` or `render_png_for(sheet, &viewport)` with `PngTarget::Viewport { sheet, viewport }`; use `PngTarget::Range { sheet, range }` for ranges. Scale defaults to 1.0.
