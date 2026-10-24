---
"@betteroffice/pptx": minor
"@betteroffice/agents": minor
---

Adds `insertTableRow`, `deleteTableRow`, `insertTableColumn` and `deleteTableColumn`, plus matching `pptx_edit` operations. Edits keep formatting, refresh refs, undo and sync; if two people restructure one table at once, one edit wins. New cells take one line. Merged tables and cells with text fields are refused.
