---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Number block-level page and column breaks among a story's breaks instead of by their position. A break's block id changed with every edit before it, so incremental pagination saw every later page break as a changed block and could not reuse the pages after an edit: an edit near the start of a long document repaginated, rebuilt and re-encoded every later page.
