---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Renaming a sheet to a name holding a character XML cannot carry is refused as an invalid step when a formula, defined name or hyperlink references the sheet, instead of writing invalid XML; unreferenced sheets still take any name.
