---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Keeps bookmarks that sit between table rows or between cells on save, including after rows or columns are inserted or deleted in an editing session. Rust literals must set `trailing_bookmarks` on `Table`, `leading_bookmarks` on `TableCell`, and both on `TableRow`.
