---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

Runs whose character scale is 0, which Word draws at full width, now draw at full width instead of disappearing on canvas and failing the page natively. Other values outside 1 to 600 are ignored or held to 600.
