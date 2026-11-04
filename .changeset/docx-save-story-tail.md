---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Editing a document that ends with a page break no longer rewrites the whole body on save, so untouched paragraphs keep their tracked changes, proofing marks and other formatting.
