`workbook-0.2.1-arrays-untouched.update.bin` and `workbook-0.2.1-arrays-edited.update.bin` were produced with the `betteroffice-xlsx` crate at release cf3d220f7 (`@betteroffice/xlsx@0.2.1`, collaboration schema 6) from `dynamic-arrays.xlsx`, a dynamic `SORT` over C1:C3 and a legacy array over E1:E2:

```rust
let untouched = Workbook::open_collaborative(&source, 7_001)?;
let snapshot = untouched.encode_state_as_update_v1();

let mut edited = Workbook::open_collaborative(&source, 7_002)?;
edited.edit_cell(SheetId(0), CellRef::parse_a1("G1")?, "note", CalculationOptions::default())?;
let snapshot = edited.encode_state_as_update_v1();
```

Schema 6 stores the arrays' cached results as constants beside plain anchor formulas.

`workbook-b153acd5b-arrays-anchors-edited.update.bin` was produced with the `betteroffice-xlsx` crate at b153acd5b (main before collaboration schema 7, schema 6) from the same workbook, editing both array anchors in a collaborative replica:

```rust
let mut edited = Workbook::open_collaborative_recalculated(&source, 7_402, CalculationOptions::default())?;
edited.edit_cell(SheetId(0), CellRef::parse_a1("C1")?, "=_xlfn._xlws.SORT(A1:A3,1,-1)", options)?;
edited.edit_cell(SheetId(0), CellRef::parse_a1("E1")?, "=A1:A2*3", options)?;
let snapshot = edited.encode_state_as_update_v1();
```

That release drops an anchor's array once its formula is edited collaboratively, on the editing replica and on every replica restoring the snapshot, so both anchors are plain formulas in it.
