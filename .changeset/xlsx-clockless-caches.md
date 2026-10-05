---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation without a `nowSerial` clock keeps the saved results of `TODAY()` and `NOW()` formulas and of formulas that depend on them.
