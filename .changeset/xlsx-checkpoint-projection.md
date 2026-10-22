---
"@betteroffice/xlsx": patch
"@betteroffice/docx": patch
---

Workbooks with sized hidden rows or columns checkpoint their recovery state again, and recovery restores undo history after row inserts above row default styles. A checkpoint that cannot be captured is logged and retried while the session keeps working and keeps its recovery log.
