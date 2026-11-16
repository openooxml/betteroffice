---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": minor
---

XLSX rows keep their authored heights under an unusual default row height. Wrapped text follows Excel's line spacing for Arial-like fonts and shows only the lines that fit its whole row, at any zoom and scroll position; the display list's `TextWrap` gains `fitHeight`.
