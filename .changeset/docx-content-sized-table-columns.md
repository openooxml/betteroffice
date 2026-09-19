---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Size a table's unpriced columns from their content. When a table states no width of its own, Word treats `w:gridCol` as a hint and re-measures every column no `w:tcW` prices, so a stored grid goes stale as soon as the content changes. Those columns now widen to their widest unwrapped cell content, spending only the room left inside the table's budget and sharing it in proportion to the demands when it cannot cover them all. Columns never shrink, so a cell can only wrap onto fewer lines; tables that state a width, and columns a cell prices, keep the declared grid exactly as before.
