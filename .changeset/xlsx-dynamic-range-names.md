---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Rows and columns can be inserted and deleted again on a sheet a defined name points at with the dynamic-range idiom — a whole-column reference inside a call such as `SUM(Data!$A:$A)`, the range operator applied to `INDEX`, or the `Data!#REF!` Excel leaves behind. Those names now move with the grid; an edit that would genuinely strand a name is still refused.
