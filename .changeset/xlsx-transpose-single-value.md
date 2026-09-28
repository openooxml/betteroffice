---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Recognise `TRANSPOSE` in the formula engine. It swaps the rows and columns of a range, array or scalar, with a blank cell transposing to `0`; errors propagate.
