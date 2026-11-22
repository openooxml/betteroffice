---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/rust-crates': minor
---

Row and column edits that reshape a range plotted through an unrebuildable cache or would split two sheets' shared chart part are refused upfront. Previously the next save failed. Cell edits under shifted ranges save. Drop the final original-workbook and sheet tuple from `chart_bytes`.
