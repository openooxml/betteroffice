---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Adds `sheets[].visibility` to `readCells` results by default, including `unknown` when visibility cannot be determined. Exports the `XlsxSheetVisibility` type from both package roots.
