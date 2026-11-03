---
"@betteroffice/xlsx-native": minor
"@betteroffice/xlsx": patch
---

Gives the native xlsx engine the same 4 GiB memory budget per thread as the WebAssembly engine. A call that takes the engine's live allocations past the budget fails: worker sessions fail with the typed out-of-memory error, and a workbook handle refuses every later call except dispose().
