`workbook-0.2.1-arrays-untouched.update.bin` and `workbook-0.2.1-arrays-edited.update.bin` were produced with the `betteroffice-xlsx` crate at release cf3d220f7 (`@betteroffice/xlsx@0.2.1`, collaboration schema 6) from `dynamic-arrays.xlsx`, a dynamic `SORT` over C1:C3 and a legacy array over E1:E2:

```rust
let untouched = Workbook::open_collaborative(&source, 7_001)?;
let snapshot = untouched.encode_state_as_update_v1();

let mut edited = Workbook::open_collaborative(&source, 7_002)?;
edited.edit_cell(SheetId(0), CellRef::parse_a1("G1")?, "note", CalculationOptions::default())?;
let snapshot = edited.encode_state_as_update_v1();
```

Schema 6 stores the arrays' cached results as constants beside plain anchor formulas.
