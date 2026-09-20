---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Add `TANH` to the XLSX formula engine. Workbooks calling it read `#NAME?` today because the name does not resolve; it now returns the hyperbolic tangent, coercing its argument like the other one-argument math functions (blank → 0, `TRUE` → 1, non-numeric text → `#VALUE!`).
