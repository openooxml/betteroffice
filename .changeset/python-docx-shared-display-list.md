---
'@betteroffice/python-docx': patch
---

`Layout.display_list` returns the same `DisplayList` on every access instead of a fresh copy, and a display list builds its JSON text only when `json`, `to_dict` or `write` asks for it, so a large layout holds far less memory.
