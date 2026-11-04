---
'@betteroffice/python-vsdx': minor
---

Adds opt-in PNG export through `render_png({"page": 0}, scale=1.0)`. Use a zero-based page target and read `Png.bytes`, dimensions and skipped images, or call `Png.write(path)`.
