---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation without a `nowSerial` clock keeps the saved results of `TODAY()` and `NOW()` formulas instead of writing `#VALUE!`; Excel recalculates them on open.
