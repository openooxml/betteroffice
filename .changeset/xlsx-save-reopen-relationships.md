---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A workbook opened from a package with duplicate workbook relationship ids, or with a sheet stored on its calc chain part, now saves a file that opens again. Sheets are read only from the workbook's own sheet list, the one a save rewrites.
