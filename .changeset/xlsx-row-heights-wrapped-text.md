---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": minor
---

XLSX wrapped text shows only the lines whose top half fits its whole row or merged area, at any zoom and scroll position; the display list's `TextWrap` gains `maxLines`, which every painter obeys.
