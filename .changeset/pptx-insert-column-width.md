---
"@betteroffice/pptx": patch
"@betteroffice/agents": minor
---

Adds opt-in `width: "copy"` to `insert_table_column` and `insertTableColumn` to preserve existing column widths and refit the table frame. Omitting `width` or choosing `"split"` keeps the current behavior.
