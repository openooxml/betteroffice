---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': patch
'@betteroffice/rust-crates': minor
'@betteroffice/python-xlsx': minor
---

`replaceText` takes `within: 'workbook'`, which skips hidden, very hidden and non-worksheet sheets, so replacing across all sheets leaves hidden sheets alone.
