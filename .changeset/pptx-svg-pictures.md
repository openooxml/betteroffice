---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': minor
'@betteroffice/python-pptx': minor
'@betteroffice/rust-crates': minor
---

Render PPTX pictures stored as SVG in both backends. The browser packages decode them with the browser's own SVG support, and the Rust and Python `render_png` paths rasterize them natively: shapes, paths, fills, strokes, gradients, clip paths, `<use>` and class-based stylesheets draw, under a sandbox that budgets the cost of expanding and painting a document before any of it runs. A tiled SVG repeats at its intrinsic size.
