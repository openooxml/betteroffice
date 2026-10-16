---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": patch
---

`openWorkbook` accepts a `calculation` option with a `WorkbookCalculationContext` (`nowSerial`, `randSeed`), and `WorkbookHandle.setCalculationContext` changes it, so date, time and random functions such as `RANDBETWEEN` compute the same values on every handle that shares the context.
