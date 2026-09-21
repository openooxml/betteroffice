---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation now supports the trigonometric family (`SIN`, `COS`, `TAN`, `SINH`, `COSH`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`), linear-fit statistics (`CORREL`, `SLOPE`, `INTERCEPT`, `COVARIANCE.P`, `COVARIANCE.S`), order statistics (`PERCENTILE.INC`, `QUARTILE.INC`), week numbering (`WEEKNUM` including the ISO type 21, `ISOWEEKNUM`) and delimiter splitting (`TEXTBEFORE`, `TEXTAFTER` with instance counting from either end, case-insensitive matching and an if-not-found fallback). On the 165-workbook calculation benchmark this takes `sheetpedia-441e60078db9` from 306 to 879 of 918 cells and `sheetpedia-8dc99d663d40` from 12 to 380 of 636, moving the corpus from 86.4906% to 87.3023% of cells matching Excel with no workbook regressing.
