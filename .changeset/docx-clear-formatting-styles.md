---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Clear Formatting in the DOCX editor returns text to what its paragraph and table styles give, so headings stay bold. `formatText` clear `'all'` no longer changes a run's text direction.
