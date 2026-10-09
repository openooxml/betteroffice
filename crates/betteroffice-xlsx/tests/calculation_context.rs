use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellState, CellValue, EditRequest, Op,
    Sheet, SheetId, Workbook, WorkbookModel,
};
use serde_json::json;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for name in ["Data", "Other"] {
        let mut sheet = Sheet::new(name);
        for (address, formula) in [
            ("A1", "NOW()"),
            ("A2", "TODAY()"),
            ("A3", "RANDBETWEEN(1,1000000)"),
            ("A4", "RANDBETWEEN(1,1000000)"),
            ("A5", "RANDBETWEEN(1,1000000)*1000+RANDBETWEEN(1,1000000)"),
            ("B3", "RANDBETWEEN(1,1000000)"),
            ("C3", "MAKEARRAY(2,2,LAMBDA(r,c,RANDBETWEEN(1,1000000)))"),
        ] {
            sheet.set_cell(
                cell(address),
                Cell {
                    value: CellValue::Number { value: -1.0 },
                    formula: Some(formula.to_owned()),
                    ..Cell::default()
                },
            );
        }
        sheet.set_array_formula(cell("C3"), CellRange::parse_a1("C3:D4").unwrap());
        model.sheets.push(sheet);
    }
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn context() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(45_000.75),
    }
}

fn value(workbook: &Workbook, sheet: u32, address: &str) -> CellValue {
    workbook
        .sheet(SheetId(sheet))
        .unwrap()
        .cell(cell(address))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn values(workbook: &Workbook) -> Vec<CellValue> {
    (0..2)
        .flat_map(|sheet| {
            [
                "A1", "A2", "A3", "A4", "A5", "B3", "C3", "C4", "D3", "D4", "B10", "B11", "B12",
                "C10", "D10",
            ]
            .map(|address| value(workbook, sheet, address))
        })
        .collect()
}

fn replay(workbook: &mut Workbook, options: CalculationOptions) -> Vec<(Vec<CellValue>, Vec<u8>)> {
    let mut states = vec![(values(workbook), workbook.save().unwrap())];
    workbook
        .edit_cell(
            SheetId(0),
            cell("B10"),
            "=NOW()+RANDBETWEEN(1,1000000)",
            options,
        )
        .unwrap();
    states.push((values(workbook), workbook.save().unwrap()));
    workbook
        .edit_cells(
            SheetId(1),
            &[
                "=TODAY()",
                "=RANDBETWEEN(1,1000000)",
                "=RANDBETWEEN(1,1000000)",
            ]
            .into_iter()
            .enumerate()
            .map(|(index, input)| CellInput {
                cell: CellRef::new(9 + index as u32, 1),
                input: input.to_owned(),
            })
            .collect::<Vec<_>>(),
            options,
        )
        .unwrap();
    states.push((values(workbook), workbook.save().unwrap()));
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "calculation": { "nowSerial": options.now_serial },
        "steps": [{
            "op": "setFormulas",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "C10:D10" } },
            "formulas": [["NOW()+RANDBETWEEN(1,1000000)", "TODAY()+RANDBETWEEN(1,1000000)"]]
        }]
    }))
    .unwrap();
    assert!(workbook.apply_edits(&request).unwrap().unwrap().applied);
    states.push((values(workbook), workbook.save().unwrap()));
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(1),
                at: cell("C10"),
                cell: CellState {
                    formula: Some(
                        "NOW()+TODAY()+RANDBETWEEN(1,1000000)+RANDBETWEEN(1,1000000)".to_owned(),
                    ),
                    ..CellState::default()
                },
            }],
            options,
        )
        .unwrap();
    states.push((values(workbook), workbook.save().unwrap()));
    assert!(workbook.undo(options).unwrap().applied);
    states.push((values(workbook), workbook.save().unwrap()));
    assert!(workbook.redo(options).unwrap().applied);
    states.push((values(workbook), workbook.save().unwrap()));
    states
}

#[test]
fn identical_context_and_calls_produce_identical_values_and_saved_bytes() {
    let bytes = fixture();
    let options = context();
    let mut first =
        Workbook::open_recalculated_with_seed(&bytes, options, Some(0x1234_5678)).unwrap();
    let mut second =
        Workbook::open_recalculated_with_seed(&bytes, options, Some(0x1234_5678)).unwrap();
    assert_eq!(
        value(&first, 0, "A1"),
        CellValue::Number { value: 45_000.75 }
    );
    assert_eq!(
        value(&first, 0, "A2"),
        CellValue::Number { value: 45_000.0 }
    );
    assert_ne!(value(&first, 0, "A3"), value(&first, 1, "A3"));
    assert_ne!(value(&first, 0, "A3"), value(&first, 0, "B3"));
    assert_eq!(replay(&mut first, options), replay(&mut second, options));
}

#[test]
fn changing_only_the_seed_changes_random_values() {
    let bytes = fixture();
    let first = Workbook::open_recalculated_with_seed(&bytes, context(), Some(0)).unwrap();
    let second = Workbook::open_recalculated_with_seed(&bytes, context(), Some(u32::MAX)).unwrap();
    assert_eq!(value(&first, 0, "A1"), value(&second, 0, "A1"));
    assert_eq!(value(&first, 0, "A2"), value(&second, 0, "A2"));
    assert_ne!(value(&first, 0, "A3"), value(&second, 0, "A3"));
    assert_ne!(value(&first, 0, "A4"), value(&second, 0, "A4"));
}

#[test]
fn seeded_random_values_survive_full_and_incremental_recalculation() {
    let options = context();
    let mut workbook =
        Workbook::open_recalculated_with_seed(&fixture(), options, Some(42)).unwrap();
    let before = values(&workbook);
    assert!(workbook.recalculate_all(options).changed.is_empty());
    assert_eq!(values(&workbook), before);
    workbook
        .edit_cell(SheetId(0), cell("Z99"), "10", options)
        .unwrap();
    assert_eq!(values(&workbook), before);
    workbook
        .edit_cell(
            SheetId(0),
            cell("C1"),
            "=A3+RANDBETWEEN(1,1000000)",
            options,
        )
        .unwrap();
    assert_eq!(values(&workbook), before);
    workbook.recalculate_all(options);
    assert_eq!(values(&workbook), before);
}

#[test]
fn earlier_independent_random_cells_preserve_values_on_full_and_incremental_recalculation() {
    let options = context();
    let mut workbook =
        Workbook::open_recalculated_with_seed(&fixture(), options, Some(42)).unwrap();
    for sheet in 0..2 {
        for address in ["A3", "A4", "B3", "C3", "C4", "D3", "D4"] {
            assert!(
                matches!(value(&workbook, sheet, address), CellValue::Number { value } if (1.0..=1_000_000.0).contains(&value))
            );
        }
        assert!(
            matches!(value(&workbook, sheet, "A5"), CellValue::Number { value } if (1_001.0..=1_001_000_000.0).contains(&value))
        );
        assert_eq!(
            workbook
                .sheet(SheetId(sheet))
                .unwrap()
                .array_formula(cell("C3")),
            Some(CellRange::parse_a1("C3:D4").unwrap())
        );
        assert_ne!(value(&workbook, sheet, "C3"), value(&workbook, sheet, "C4"));
        assert_ne!(value(&workbook, sheet, "C3"), value(&workbook, sheet, "D3"));
    }
    let before = values(&workbook);
    workbook
        .edit_cell(SheetId(0), cell("B1"), "=RANDBETWEEN(1,1000000)", options)
        .unwrap();
    assert!(
        matches!(value(&workbook, 0, "B1"), CellValue::Number { value } if (1.0..=1_000_000.0).contains(&value))
    );
    assert_eq!(values(&workbook), before);
    assert!(workbook.recalculate_all(options).changed.is_empty());
    assert_eq!(values(&workbook), before);
}

#[test]
fn sharing_only_the_clock_keeps_random_draws_unpinned() {
    let mut workbook = Workbook::open_recalculated(&fixture(), context()).unwrap();
    let before = value(&workbook, 0, "A3");
    assert!(matches!(before, CellValue::Number { value } if (1.0..=1_000_000.0).contains(&value)));
    assert_eq!(workbook.rand_seed(), None);
    let mut draws = Vec::new();
    for input in ["10", "20", "30", "40"] {
        workbook
            .edit_cell(SheetId(0), cell("Z99"), input, context())
            .unwrap();
        assert_eq!(
            value(&workbook, 0, "A1"),
            CellValue::Number { value: 45_000.75 }
        );
        draws.push(value(&workbook, 0, "A3"));
    }
    assert!(draws.iter().any(|draw| *draw != before));
}

#[test]
fn workbook_seed_can_be_pinned_changed_and_cleared() {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open_recalculated(&fixture(), options).unwrap();
    assert_eq!(workbook.rand_seed(), None);
    workbook.set_rand_seed(Some(42));
    assert_eq!(workbook.rand_seed(), Some(42));
    workbook.recalculate_all(options);
    let pinned = value(&workbook, 0, "A3");
    workbook.set_rand_seed(Some(7));
    workbook.recalculate_all(options);
    assert_ne!(value(&workbook, 0, "A3"), pinned);
    workbook.set_rand_seed(Some(42));
    workbook
        .edit_cell(SheetId(0), cell("Z99"), "10", options)
        .unwrap();
    assert_eq!(value(&workbook, 0, "A3"), pinned);
    workbook.set_rand_seed(None);
    assert_eq!(workbook.rand_seed(), None);
    let mut draws = Vec::new();
    for _ in 0..4 {
        workbook.recalculate_all(options);
        draws.push(value(&workbook, 0, "A3"));
    }
    assert!(draws.iter().any(|draw| *draw != pinned));
}

#[test]
fn batch_requests_override_only_the_clock_when_present() {
    let mut workbook =
        Workbook::open_recalculated_with_seed(&fixture(), context(), Some(42)).unwrap();
    let random = value(&workbook, 0, "A3");
    for (index, calculation) in [
        None,
        Some(json!({})),
        Some(json!({ "nowSerial": 46_000.25 })),
    ]
    .into_iter()
    .enumerate()
    {
        let mut request = json!({
            "expectVersion": workbook.version(),
            "steps": [{
                "op": "setCellInputs",
                "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "Z99" } },
                "inputs": [[index.to_string()]]
            }]
        });
        if let Some(calculation) = calculation {
            request["calculation"] = calculation;
        }
        let result: serde_json::Value = serde_json::from_str(
            &workbook
                .apply_edits_json_with_calculation(&request.to_string(), context())
                .unwrap(),
        )
        .unwrap();
        assert_eq!(result["ok"], true);
        let expected = if index < 2 { 45_000.75 } else { 46_000.25 };
        assert_eq!(
            value(&workbook, 0, "A1"),
            CellValue::Number { value: expected }
        );
        assert_eq!(value(&workbook, 0, "A3"), random);
        assert_eq!(workbook.rand_seed(), Some(42));
    }
}

#[test]
fn clock_only_batch_overrides_preserve_seeded_values_and_saved_bytes() {
    let bytes = fixture();
    let mut first = Workbook::open_recalculated_with_seed(&bytes, context(), Some(42)).unwrap();
    let mut second = Workbook::open_recalculated_with_seed(&bytes, context(), Some(42)).unwrap();
    let random = value(&first, 0, "A3");
    for workbook in [&mut first, &mut second] {
        let request = json!({
            "expectVersion": workbook.version(),
            "calculation": { "nowSerial": 46_000.25 },
            "steps": [{
                "op": "setFormulas",
                "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "B10" } },
                "formulas": [["RANDBETWEEN(1,1000000)"]]
            }]
        });
        let result: serde_json::Value = serde_json::from_str(
            &workbook
                .apply_edits_json_with_calculation(&request.to_string(), context())
                .unwrap(),
        )
        .unwrap();
        assert_eq!(result["ok"], true);
        assert_eq!(result["applied"], true);
        assert_eq!(workbook.rand_seed(), Some(42));
        assert_eq!(
            value(workbook, 0, "A1"),
            CellValue::Number { value: 46_000.25 }
        );
        assert_eq!(value(workbook, 0, "A3"), random);
        assert!(
            matches!(value(workbook, 0, "B10"), CellValue::Number { value } if (1.0..=1_000_000.0).contains(&value))
        );
    }
    assert_eq!(values(&first), values(&second));
    assert_eq!(first.save().unwrap(), second.save().unwrap());
}
