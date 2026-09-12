`table-basic.pptx` is a synthetic, one-slide repro built from the same minimal
package skeleton as the other fixtures here. It contains no private corpus
material.

One 3x3 table in a `p:graphicFrame`, sized 5715000 x 1112520 EMU, carrying every
feature the table work has to handle:

- `a:tblGrid` with three unequal columns (1828800, 1219200, 2667000 EMU) and an
  `a:tr@h` on every row, so column offsets and row heights are both observable.
- A `gridSpan="2"` origin in row 0 followed by its `hMerge="1"` continuation, and
  a `rowSpan="2"` origin in row 1 followed by its `vMerge="1"` continuation. A
  continuation carries no text and draws nothing; the origin spans.
- An explicit `a:tcPr` `solidFill` on the `rowSpan` origin and an explicit
  `a:lnB` on one cell, so direct cell formatting can be told apart from what the
  table style contributes.
- `a:tblPr` `firstRow="1" bandRow="1"` with a `tableStyleId` that resolves in
  `ppt/tableStyles.xml`, which defines `wholeTbl` borders and fill, a `band1H`
  fill and a `firstRow` fill with bold white text.

The style part is deliberately resolvable: 56 of the 61 tables in the corpus
carry a `tableStyleId` and only 2 of 13 distinct ids fail to resolve, so the
resolving case is the one worth pinning first.

Today the frame renders as a placeholder, and `layout.rs` painted the first cell
as loose slide text on top of it. See the [DrawingML table
reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table).
