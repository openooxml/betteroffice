---
'@betteroffice/python-xlsx': minor
---

Changes `render_png` to accept `{"sheet": 0, "range": "A1:B2"}` or `{"sheet": 0, "viewport": {...}}`, with keyword-only `scale=1.0`. Replace sheet names and range keywords with zero-based target dictionaries; dimension caps are no longer arguments.
