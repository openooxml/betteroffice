---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Add `MMULT` to the XLSX formula engine. The matrix product requires `cols(array1)` to equal `rows(array2)`, gives `#VALUE!` for a dimension mismatch or any empty, text or logical operand cell, and propagates an argument's error. `MMULT` returns the whole product as an array, which spills from a dynamic-array formula and fills a legacy array formula's range.
