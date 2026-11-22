---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

After `WorkbookHandle.recalculate()`, saved files carry the recalculated values; formulas the engine cannot evaluate keep the result the file stored. The recalculation at open does not change what a save without edits writes.
