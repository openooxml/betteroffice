---
"@betteroffice/docx": patch
---

Adds table structure commands to the worker-only DOCX editor: `insertTableRow`, `insertTableColumn`, `deleteTableRow`, `deleteTableColumn`, `deleteTable`, `mergeTableCells`, `splitTableCell` and `setTableCellShading`. Each command is one undo step.
