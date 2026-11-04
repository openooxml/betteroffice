---
'@betteroffice/python-xlsx': minor
---

`render_png` takes a target: `{"sheet": name_or_index, "range": ...}` or `{"sheet": name_or_index, "viewport": {...}}`, with `scale`, `max_width` and `max_height` as keywords. `range` is an A1 string or an address dictionary. For the old `render_png(sheet)`, which frames the used range, call `render_sheet(sheet)`.
