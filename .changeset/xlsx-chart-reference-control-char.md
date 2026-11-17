---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

A workbook whose chart reference holds a character XML cannot carry now opens, and its chart is kept as authored. An edit that would have to rewrite such a reference is refused up front, so saving never fails.
