---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Date functions in workbooks that use the 1904 date system now return 1904 serials: `DATE(2026,11,1)` is 44865, as in Excel. `nowSerial` is always a 1900-system serial, so TODAY(), NOW() and typed dates without a year match the calendar in 1904 workbooks too.
