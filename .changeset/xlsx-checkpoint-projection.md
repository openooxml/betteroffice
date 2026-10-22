---
"@betteroffice/xlsx": patch
"@betteroffice/docx": patch
---

Workbooks with sized hidden rows or columns checkpoint their recovery state again, and recovery keeps undo history after row inserts above row default styles. A checkpoint that cannot be captured is retried while the session keeps working.
