---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: a sheet declaring `sheetFormatPr/@zeroHeight` now hides the rows that carry no height of their own, as Excel does. The attribute was read nowhere, so those rows took a fitted or default height and rendered visible.
