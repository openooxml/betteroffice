---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation now supports `LINEST` over one or more predictors, with the full statistics block its fourth argument asks for, and a scalar cell reading an array-only result through `INDEX` now evaluates it as an array instead of returning `#VALUE!`. The fit uses Householder QR rather than the normal equations, so a polynomial design such as `LINEST(y, x^{1,2,3,4,5,6})` stays conditioned. On the 165-workbook calculation benchmark this takes `sheetpedia-def4407b7c16` from 0 to 1989 of 1989 cells and `sheetpedia-c70e22c95a64` from 36 to 93 of 97, moving the corpus from 88.4700% to 90.1053% of cells matching Excel with no workbook regressing.
