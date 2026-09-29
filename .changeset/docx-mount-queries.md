---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Selection and story queries on large documents no longer trigger whole-document rework; the first read of every story after an edit was quadratic in story count.
