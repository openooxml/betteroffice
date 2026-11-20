---
'@betteroffice/python-xlsx': minor
---

`Sheet.render_png()` renders its own sheet again when called without a target, with `scale`, `range`, `max_width` and `max_height` as before. A target without a `sheet` now renders that sheet too.
