---
'@betteroffice/python-xlsx': patch
---

Preserve Python strings as text in cell assignments, batches, and proposals so numeric strings retain their decimal places and leading zeros. Formula strings and apostrophe escapes remain supported; use Python numbers and booleans to write those types, and `None` to clear a cell.
