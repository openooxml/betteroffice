---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
---

Rich text written by Excel keeps its runs in `Workbook::cell_text_runs`: font family, scheme or charset, other run properties and phonetic guides no longer make the string plain. Replacing text in such a cell now says it holds rich text.
