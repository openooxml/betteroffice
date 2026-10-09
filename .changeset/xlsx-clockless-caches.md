---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation without a `nowSerial` clock no longer overwrites the saved results of `TODAY()`, `NOW()` and year-less `DATEVALUE()` formulas with `#VALUE!`. Passing `nowSerial` recalculates them normally.
