---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: `RANK`, `RANK.EQ`, `RANK.AVG` and `MATCH` given a range of values now answer one result per value instead of a single error, which is the shape the classic `LOOKUP(2, 1/(RANK(col,col,1)=n), col)` sorting idiom depends on. Their reference argument stays a range. On the 165-workbook calculation benchmark this takes `sheetpedia-eb5de1652388` from 133 to 945 of 946 cells, moving the corpus from 90.1053% to 90.7508% with no workbook regressing.
