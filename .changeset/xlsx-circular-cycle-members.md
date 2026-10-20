---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Recalculation flags only the formulas on a real reference cycle as circular; their dependents and formulas whose ranges only look circular now show computed values.
