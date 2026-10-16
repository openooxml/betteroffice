---
"@betteroffice/docx": patch
---

Saving an edited paragraph keeps its editing permission ranges in place. When an edit replaces the text a permission marker sits in, the save now fails with an error naming the part instead of dropping the marker.
