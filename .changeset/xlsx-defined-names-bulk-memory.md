---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Editing many defined names in one `applyOps` call now takes time and memory linear in the number of names. Re-committing all 66,000 names of a large workbook used to exhaust memory.
