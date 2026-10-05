---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation without a `nowSerial` clock keeps the saved results of cells that use `TODAY()`, `NOW()` or a year-less `DATEVALUE()` and recalculates every other cell. Passing `nowSerial` recalculates them normally.
