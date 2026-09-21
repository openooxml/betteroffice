---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation now supports `QUOTIENT`, `SERIESSUM`, `ISEVEN`, `ISODD` and `INDIRECT`. `INDIRECT` resolves as a reference as well as a value, so it works as the argument to `SUM`, `INDEX` and the other range-taking functions; it accepts only a reference-shaped string, so the text cannot smuggle in another call. On the 165-workbook calculation benchmark this takes `sheetpedia-9b13ef8593b1` from 0 to 2886 of 2916 cells and `sheetpedia-58e0d15358b7` from 0 to 28 of 28, moving the corpus from 68.4866% to 70.8031% of cells matching Excel with no workbook regressing.
