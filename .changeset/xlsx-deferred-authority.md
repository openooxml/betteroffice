---
'@betteroffice/rust-crates': minor
'@betteroffice/xlsx': patch
'@betteroffice/python-xlsx': patch
---

Opening an XLSX workbook defers its collaboration state until an edit, encode, snapshot or recovery checkpoint needs it, so reading, rendering, recalculating and saving a large workbook take about 70% less memory. New `try_client_id` and `try_encode_state_*` methods return errors instead of panicking.
