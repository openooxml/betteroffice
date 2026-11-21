`leading-equals.xlsx` holds formulas stored with a leading `=` (`<f>=A1*2</f>`, `<f>=HYPERLINK(...)</f>`) and with storage prefixes (`_xlfn.CONCAT`, `=_xlfn.XLOOKUP`). The two `.bin` files were produced from it with the crates at commit `f92d5602`, which read those formulas as stored:

```rust
let mut writer = Workbook::open_collaborative(&source, 7_101)?;
writer.edit_cell(SheetId(0), CellRef::parse_a1("B1")?, "edited", CalculationOptions::default())?;
let update = writer.encode_state_as_update_v1();
let checkpoint = writer.recovery_checkpoint(0, None)?;
```

The checkpoint restores only with the engine version that wrote it (`0.5.1`).
