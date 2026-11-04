---
'@betteroffice/rust-crates': minor
---

Changes `render_png` to take `PngTarget { slide }` and `&RenderOptions`. Replace the positional slide index with a zero-based target; background, shadow limits and default scale are unchanged.
