---
'@betteroffice/python-docx': minor
---

Changes `render_png` to accept `{"page": 0}` with keyword-only `scale=1.0`. Migrate display-list callers to `render_display_list_png(display_list, page)`; page indexes remain zero-based.
