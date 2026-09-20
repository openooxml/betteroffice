---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: row heights stored against a cached `sheetFormatPr/@defaultRowHeight` now render at the height Excel recomputes for the sheet's normal font
