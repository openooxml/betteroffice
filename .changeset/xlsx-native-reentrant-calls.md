---
"@betteroffice/xlsx-native": patch
---

A native XLSX callback that calls a method on, or frees, a workbook or snapshot builder that is already in use now gets an error instead of reaching it mid-call, as the wasm binding does. An error thrown by a progress callback no longer surfaces from the call it ran in.
