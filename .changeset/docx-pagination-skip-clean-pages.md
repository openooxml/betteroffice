---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Changes far apart in one layout pass, such as a round of host proposals, now lay out again only the pages around each change. `layout_document_incremental_ranges` reports the page ranges placed afresh.
