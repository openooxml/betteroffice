---
'@betteroffice/rust-crates': minor
---

Replaces document PNG export with `render_png(PageTarget { page }, &RenderOptions)`. Migrate callers holding a display list to `render_display_list_png(&display_list, page)`; page indexes remain zero-based and scale defaults to 1.0.
