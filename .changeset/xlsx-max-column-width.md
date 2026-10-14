---
"@betteroffice/xlsx": patch
---

Opens workbooks with a column wider than Excel's 255-character limit, which previously failed with "workbook contains an invalid column width". The column renders at 255 and the authored width is kept on save.
