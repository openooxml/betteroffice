---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Answer selection and story queries on large documents without whole-document rework. Per-story segment and chunk caches no longer rescan every cached story on each insert, which made the first read of every story after an edit quadratic in the story count; the toolbar's in-table check reads the cell's own parent first; font preflight collects requirements without copying every lowered story; and story membership checks no longer list and sort every story id.
