---
'@betteroffice/python-pptx': minor
---

Changes `render_png` to accept `{"slide": 0}` with keyword-only scale, background and shadow options. Replace positional slide indexes with zero-based targets; translate slide IDs to indexes before rendering.
