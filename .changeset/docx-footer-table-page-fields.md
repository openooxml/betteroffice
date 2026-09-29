---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Show the full page number when a header or footer puts its PAGE or NUMPAGES field in a table. Only top-level header and footer paragraphs got per-page field widths before, so a field in a table cell kept the width of its cached result and every digit past it was clipped at the cell edge: a right-aligned `Page {PAGE}` cached as "2" showed "Page 2" on page 26 and "Page 1" on page 122.
