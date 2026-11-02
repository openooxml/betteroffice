---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

After a recalculation, saved files carry the recalculated values: the new `WorkbookHandle.recalculate()`, or opening a file that asks for a full calculation on load. Formulas the engine cannot evaluate keep the result the file stored. Without either, saved values stay byte for byte.
