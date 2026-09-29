---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Load display-list query pages without rewriting the whole page array. A page load used to send the store a reuse entry for every other page, and the store rebuilt its page array on every load, so querying each page of a document once cost time quadratic in its page count. A page load now replaces only the loaded pages and keeps every other page in place: one range query per page drops from about 650-850 ms to about 150 ms at 2,000 pages, and from about 200 ms to 60-115 ms at 1,000.
