---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx-native": patch
---

Native XLSX rendering shapes and draws cell text through the shared font stack and reuses glyph outlines across a sheet, about 10% faster. Pixel output and text measurement are unchanged.
