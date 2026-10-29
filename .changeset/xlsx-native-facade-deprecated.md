---
'@betteroffice/xlsx-native': patch
---

Deprecates the asynchronous `openWorkbook` and `Workbook` facade in favour of `@betteroffice/xlsx`, which automatically uses this package as its native engine when installed. The facade is removed in 0.7.
