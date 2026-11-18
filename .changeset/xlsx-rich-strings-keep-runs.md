---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
---

Rich text written by Excel keeps its runs in `Workbook::cell_text_runs`. Font family, scheme and charset, run properties the reader does not model, and phonetic guides no longer turn the whole string into plain text. Replacing text in such a cell now says it holds rich text, instead of that it may.
