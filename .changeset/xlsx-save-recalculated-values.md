---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

After `WorkbookHandle.recalculate()`, saved files carry the recalculated values; formulas the engine cannot evaluate keep the result the file stored. A save without edits, showing the stored active sheet, does not write the recalculation at open.
