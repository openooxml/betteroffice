---
"@betteroffice/xlsx-native": patch
---

Native XLSX callbacks now get an error if they call methods on a workbook or snapshot builder already in use. Errors thrown by progress callbacks no longer interrupt workbook operations.
