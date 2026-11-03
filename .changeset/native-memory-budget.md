---
"@betteroffice/xlsx-native": patch
"@betteroffice/xlsx": patch
---

The native xlsx engine has the same 4 GiB memory budget per thread as the WebAssembly engine. A call past the budget fails the worker session with the typed out-of-memory error, and a failed workbook handle refuses later calls except dispose().
