---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Reserve body room on each page for the header and footer that page shows, as Word does, so a tall first-page or never-shown even-page band no longer shortens every page. Even pages are now chosen by their displayed page number, and a missing even band under `w:evenAndOddHeaders` stays blank instead of repeating the default one.
