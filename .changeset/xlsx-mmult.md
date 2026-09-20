---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Add `MMULT` to the XLSX formula engine. The matrix product requires `cols(array1)` to equal `rows(array2)`, gives `#VALUE!` for a dimension mismatch or any empty, text or logical operand cell, and propagates an argument's error. A cell holds a single value and the model records no array-formula range, so `MMULT` returns the top-left element of its product — the value Excel caches in the anchor cell of the array formula that entered it, while the remaining cells of a legacy CSE range carry no formula and keep their stored values.
