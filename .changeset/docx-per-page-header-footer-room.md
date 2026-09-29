---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Reserve body room on each page for the header and footer that page shows, as Word does. A section's pages all reserved its tallest header and footer, so a tall first-page footer under `w:titlePg`, or an even-page header without `w:evenAndOddHeaders`, which Word never shows, shortened every page of the section. The first-page band now shortens only the first page, and the even band only even pages when even and odd headers are on.
