---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Full recalculation of an xlsx workbook is faster: formulas that did not change since the previous calculation are no longer parsed again.
