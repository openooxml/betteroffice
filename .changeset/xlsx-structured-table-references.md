---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Read `xl/tables/*.xml` into the workbook model and resolve structured table references in formulas. `Table[]`, `Table[Column]`, the `#All`/`#Headers`/`#Data`/`#Totals` bands, `[#This Row]` with its `@` shorthand, and a `[[First]:[Last]]` column span all resolve to a live area, so `VLOOKUP`, `SUMIFS`, `AVERAGEIFS` and `OFFSET` see a range rather than a copied block, and the dependency graph reads the exact rectangle each reference designates. `INDEX` now yields a reference too, so a zero or omitted index keeps a whole row or column addressable.

Measured on the 165-document Excel-rebuild calculation corpus: 90.1053% -> 91.3557% (+1,573 cells), twelve zero-scoring documents down to nine, no document regressed. Unmodified packages still re-serialize byte-identically (165/165) and recalc-save-reopen stays stable (165/165).
