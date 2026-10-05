---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

An edited save that would have to rewrite a worksheet's cells without source markup the editor does not model now fails with an error naming the sheet and the markup, and the workbook stays editable.
