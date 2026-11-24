---
"@betteroffice/agents": patch
---

`xlsx_sort` now shifts relative row references when it moves formulas. If a formula cannot be moved safely, the sort reports the affected cell before applying any edits.
