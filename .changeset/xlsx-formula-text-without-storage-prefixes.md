---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/agents': patch
'@betteroffice/rust-crates': patch
---

Shown and returned formula text drops Excel's storage prefixes (`_xlfn.`, `_xlws.`, `_xlpm.`, `_xleta.`): the formula bar, `readCells`, structured export, defined names, conditional formats, agent reads and Python `formula()` give `=SEQUENCE(2,2)`. Saved files keep the prefixes; re-committing a shown formula leaves the cell unchanged.
