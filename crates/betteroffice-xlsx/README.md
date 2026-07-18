# betteroffice-xlsx

The typed Rust API for opening, editing, calculating, rendering, and saving XLSX
workbooks.

```rust
use betteroffice_xlsx::{
    CalculationOptions, CellRef, RenderOptions, SheetId, Workbook,
};

let mut workbook = Workbook::open_recalculated(
    &xlsx_bytes,
    CalculationOptions::default(),
)?;

workbook.edit_cell(
    SheetId(0),
    CellRef::parse_a1("A1").unwrap(),
    "42",
    CalculationOptions::default(),
)?;

let png = workbook.render_sheet(SheetId(0), &RenderOptions::default())?;
let saved = workbook.save()?;
# Ok::<(), betteroffice_xlsx::Error>(())
```

Saving regenerates the package from the XLSX features represented by the model;
unsupported package parts are not retained.
The API is experimental and may change before `0.1.0`.
