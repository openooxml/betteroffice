---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
'@betteroffice/xlsx-react': minor
---

Number formats a workbook declares for built-in ids are shown as declared, as in Excel: an accounting format with a minus sign or an East Asian date no longer shows the built-in format. General stays General whatever the workbook declares for it. Shared sessions created before this version keep their previous display until the workbook is opened again. Sessions shared with clients on an older version cannot open states created by this one.
