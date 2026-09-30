---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A row of an inline table now splits across pages only where Word allows it: widow and orphan control and `w:keepLines` keep a cell paragraph's lines together, as in body text, unless the row cannot fit on a whole page.
