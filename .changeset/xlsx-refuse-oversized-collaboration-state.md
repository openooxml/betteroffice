---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Workbooks whose collaborative state exceeds its size limits are refused at open, as in 0.5.1, instead of opening and failing at the first edit.
