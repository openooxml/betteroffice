---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

`renderPng` with no `fontProvider` keeps measured metrics for symbols no face covers, so such a paragraph no longer wraps on estimated widths.
