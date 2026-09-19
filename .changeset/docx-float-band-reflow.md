---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Restart flow content below a page-anchored floating table instead of painting the table over it. A full-width `w:tblpPr` table with `w:vertAnchor="page"` and an explicit `w:tblpY` claims a band on its page; Word moves the first flow fragment reaching into that band, and everything after it, down to the band's bottom. Measured against Word 16.112.4, which restarts the flow at exactly the band bottom. The band keeps its existing handling when the shift would push content past the bottom margin.
