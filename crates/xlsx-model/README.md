# betteroffice-xlsx-model

The workbook model every other XLSX crate is written against: addresses, cell
values, styles, and the cell-access trait the calc engine reads through. Pure
data — no IO, no XML, no DOM.

- `addr` — `SheetId`, `CellRef`, `CellRange`, `RowId`, `ColId`, and the sheet
  bounds
- `value` — `CellValue` and the Excel error values
- `styles` — fonts, fills, borders, alignment, number formats, the stylesheet,
  and theme colors
- `numfmt` — number-format parsing and value formatting
- `date` — serial to calendar conversion, including the 1900 leap bug
- `workbook` — `Workbook`, `Sheet`, `Cell`, defined names, freeze panes,
  hyperlinks, and the `CellProvider` trait

`CellProvider` is the seam that keeps
[betteroffice-xlsx-calc](https://crates.io/crates/betteroffice-xlsx-calc) pure:
the formula engine reads cells exclusively through it and never knows where
they came from.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
