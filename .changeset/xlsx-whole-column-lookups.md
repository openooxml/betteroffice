---
"@betteroffice/xlsx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": patch
---

Support whole-column formula references such as `VLOOKUP(..., S:V, ...)` with limits: narrow hits evaluate without materialising the column, but wide aggregates such as `SUM(A:XFD)` return `#NUM!`, and every lookup miss scans the full column height against the shared per-recalculation budget, so a workbook with many misses can turn later formulas `#NUM!`.
