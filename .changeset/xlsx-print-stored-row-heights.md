---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Printed worksheets keep their stored row heights when the sheet's default row height is smaller than its font needs, so ranges that fit one page in Excel no longer spill onto the next.
