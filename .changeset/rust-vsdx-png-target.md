---
'@betteroffice/rust-crates': minor
---

Replaces `export_png(page, scale)` with `render_png(PageTarget { page }, &RenderOptions { scale })`. Page indexes remain zero-based and default scale is 1.0.
