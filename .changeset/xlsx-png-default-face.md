---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

XLSX text with no font of its own, including hyperlink text, now names the workbook's Normal font (or Calibri) on screen and in PNG export, so canvas no longer falls back to a system sans-serif.
