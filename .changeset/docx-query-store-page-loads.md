---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Loading display-list query pages one at a time no longer rewrites every page, so querying each page of a long document stays fast as the page count grows.
