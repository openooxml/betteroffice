---
"@betteroffice/xlsx": patch
---

Workbooks with more than 65,536 defined names open and recalculate. The name table is now bounded by 32 MiB of name and formula text, and edits that would grow it past that bound are refused without changing the workbook.
