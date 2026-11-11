---
'@betteroffice/docx': patch
---

Typing in a table cell no longer rebuilds the whole document's table layout to find the cell. The editor keeps it until the document's structure changes, which saves work per keystroke in large tables.
