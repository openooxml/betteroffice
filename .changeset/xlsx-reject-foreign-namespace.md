---
"@betteroffice/xlsx": patch
---

Opening XLSX files now rejects a workbook part with an unsupported main namespace or a non-workbook root instead of opening it, and the current document stays open.
