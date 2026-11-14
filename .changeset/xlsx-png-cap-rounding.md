---
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

PNG size caps (`max_width`, `max_height`) are no longer exceeded by 1 px at fractional scales, including in `render_sheet`.
