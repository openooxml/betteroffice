---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Clear Formatting in the DOCX editor now returns text to what its paragraph and table styles give, so a heading stays bold and saves no `w:b w:val="0"`. Character styles and direct formatting are removed; links stay.
