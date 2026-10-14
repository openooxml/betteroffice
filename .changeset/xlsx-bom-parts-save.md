---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Edited workbooks whose worksheet, workbook, styles or shared-strings part starts with a UTF-8 byte order mark now save instead of failing with "tag not closed".
