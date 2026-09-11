---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Parse `a:tbl` into a real table model: the column grid, row heights, cell spans and merge continuations, direct cell fills and borders, and the `a:tblPr` style flags. A cell's `a:tcPr` anchoring, text direction and margins fold into its text body.

Recover a table's geometry and cell formatting from a reattached source, folded into the unreleased schema 2.1.
