---
"@betteroffice/xlsx-native": patch
---

A native XLSX callback that uses or frees a workbook or snapshot builder already in use now gets an error, as with the wasm binding. A progress callback's error no longer surfaces from the call it ran in.
