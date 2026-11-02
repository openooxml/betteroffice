---
'@betteroffice/xlsx-native': patch
---

Now exposes exactly the `@betteroffice/xlsx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openWorkbook(Buffer)`/`Workbook` facade is removed.
