---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Number formats with commas after the decimal digits, such as `0.0,,"M"` and `#,##0.0,"K"`, now divide by 1,000 per comma in cell display and `TEXT()`, as Excel does.
