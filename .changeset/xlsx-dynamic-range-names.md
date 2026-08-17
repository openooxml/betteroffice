---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Rows and columns can be inserted and deleted again on a sheet a defined name points at with the dynamic-range idiom — a whole-column reference inside a call such as `SUM(Data!$A:$A)`, the range operator applied to `INDEX`, or the `Data!#REF!` Excel leaves behind. A spill (`A1#`), an implicit intersection (`@A1`) and a range written around whitespace move with the grid too. An edit whose effect on a name cannot be settled — a reference that could equally be a name of the same shape, or an unqualified one in a workbook-scope name — is still refused rather than guessed at.
