---
"@betteroffice/pptx": minor
"@betteroffice/agents": minor
---

Adds `insertTableRow`, `deleteTableRow`, `insertTableColumn` and `deleteTableColumn`, plus matching `pptx_edit` operations. Edits keep formatting, refresh refs and undo. They are refused in collaborative sessions, on merged tables and for cells with text fields. New cells take one line.
