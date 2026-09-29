---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Documents with footnotes or floating tables now relayout incrementally: an edit rebuilds only the affected pages and the float segments touching a changed block.
