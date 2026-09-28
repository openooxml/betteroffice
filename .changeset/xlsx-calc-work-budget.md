---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation keeps a hostile workbook inside its work and memory budget, so opening one returns in milliseconds with its cached values instead of running for minutes or exhausting memory: array builtins, callbacks, lifted calls, sorts, blocks of lookup keys, `SUMPRODUCT` broadcasts and array-formula fills charge their work before doing it, the text a formula builds or copies has its own byte budget per formula and per recalculation, a spill that shrinks retires only the cells its old rectangle stores, and an array formula the budget refuses keeps its rectangle and cached cells. `MAKEARRAY` charges its cells once, so a 600,000-cell block that fits the budget now evaluates instead of returning `#NUM!`.
