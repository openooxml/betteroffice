---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Open workbooks that declare a negative column width. A `<col width="-0.42">` has no extent to render, and the whole workbook was refused with `workbook contains an invalid column width` rather than the one column narrowed. A negative width now collapses to zero exactly as a hidden column already does, while the authored attribute is left alone so the worksheet part still saves byte-identically. Sample sheetpedia-91e613e07d4c went from failing to capture to 0.7294; no other sample in the 208-sample xlsx fidelity corpus changed by a single bit.
