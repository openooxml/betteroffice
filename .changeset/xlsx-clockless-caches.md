---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation without a `nowSerial` clock in a workbook using `TODAY()`, `NOW()` or `DATEVALUE()` keeps every saved formula result. Passing `nowSerial` recalculates normally.
