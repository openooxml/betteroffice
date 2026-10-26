use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellState, CellValue, ErrorValue, Op,
    Sheet, SheetId, Workbook, WorkbookModel,
};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn put(sheet: &mut Sheet, row: u32, col: u32, value: CellValue) {
    sheet.set_cell(
        CellRef::new(row, col),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn workbook(text_value: bool) -> Workbook {
    workbook_with_rows(text_value, 300)
}

fn workbook_with_rows(text_value: bool, data_rows: u32) -> Workbook {
    let mut data = Sheet::new("Data");
    for row in 1..=data_rows {
        put(&mut data, row, 0, text(&format!("a{}", row % 8)));
        put(&mut data, row, 1, text(&format!("b{}", row / 8 % 6)));
        put(&mut data, row, 2, number((row / 48 % 3) as f64));
        put(&mut data, row, 3, number((row % 17) as f64 / 8.0));
    }
    if text_value {
        put(&mut data, 250, 3, text("bad"));
    }
    let mut dashboard = Sheet::new("Dashboard");
    put(&mut dashboard, 0, 25, number(0.0));
    for col in 1..=6 {
        put(&mut dashboard, 0, col, text(&format!("b{}", col - 1)));
    }
    for row in 1..=8 {
        put(&mut dashboard, row, 0, text(&format!("a{}", row - 1)));
        for col in 1..=6 {
            let column = (b'A' + col as u8) as char;
            let end = data_rows + 1;
            dashboard.set_cell(
                CellRef::new(row, col),
                Cell {
                    formula: Some(format!(
                        "SUMPRODUCT((Data!$A$2:$A${end}=$A{})*(Data!$B$2:$B${end}={column}$1)*(Data!$C$2:$C${end}=$Z$1)*Data!$D$2:$D${end})",
                        row + 1
                    )),
                    ..Cell::default()
                },
            );
        }
    }
    let mut model = WorkbookModel::default();
    model.sheets.extend([data, dashboard]);
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    workbook
}

fn value(workbook: &Workbook, sheet: u32, row: u32, col: u32) -> CellValue {
    workbook
        .sheet(SheetId(sheet))
        .unwrap()
        .cell(CellRef::new(row, col))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn assert_dashboard(workbook: &Workbook, data_rows: u32) {
    let data: Vec<[CellValue; 4]> = (1..=data_rows)
        .map(|row| std::array::from_fn(|col| value(workbook, 0, row, col as u32)))
        .collect();
    let numeric_key = value(workbook, 1, 0, 25);
    for row in 1..=8 {
        let first_key = value(workbook, 1, row, 0);
        for col in 1..=6 {
            let second_key = value(workbook, 1, 0, col);
            let mut total = 0.0;
            let mut invalid = false;
            for values in &data {
                match &values[3] {
                    CellValue::Number { value } => {
                        if values[0] == first_key
                            && values[1] == second_key
                            && values[2] == numeric_key
                        {
                            total += value;
                        }
                    }
                    CellValue::Empty => {}
                    _ => invalid = true,
                }
            }
            let expected = if invalid {
                CellValue::Error {
                    value: ErrorValue::Value,
                }
            } else {
                number(total)
            };
            assert_eq!(value(workbook, 1, row, col), expected, "{row}, {col}");
        }
    }
}

fn edit(workbook: &mut Workbook, sheet: u32, address: &str, input: &str) {
    workbook
        .edit_cell(
            SheetId(sheet),
            CellRef::parse_a1(address).unwrap(),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

fn exercise_edits(text_value: bool) {
    let mut workbook = workbook(text_value);
    assert_dashboard(&workbook, 300);
    edit(&mut workbook, 0, "D2", "42.125");
    assert_dashboard(&workbook, 300);
    edit(&mut workbook, 0, "A2", "a2");
    assert_dashboard(&workbook, 300);
    edit(&mut workbook, 1, "A2", "a2");
    assert_dashboard(&workbook, 300);
    edit(&mut workbook, 1, "Z1", "1");
    assert_dashboard(&workbook, 300);

    let mut ops = vec![Op::InsertRows {
        sheet: SheetId(0),
        at: 100,
        count: 2,
    }];
    for row in 100..102 {
        for (col, value) in [text("a2"), text("b0"), number(1.0), number(123.5)]
            .into_iter()
            .enumerate()
        {
            ops.push(Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(row, col as u32),
                cell: CellState {
                    value,
                    ..CellState::default()
                },
            });
        }
    }
    workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap();
    assert_dashboard(&workbook, 302);
    assert!(
        workbook
            .cell(SheetId(1), CellRef::new(1, 1))
            .unwrap()
            .input
            .contains("Data!$D$2:$D$303")
    );
    workbook
        .apply_ops(
            vec![Op::DeleteRows {
                sheet: SheetId(0),
                at: 100,
                count: 3,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_dashboard(&workbook, 299);
    assert!(
        workbook
            .cell(SheetId(1), CellRef::new(1, 1))
            .unwrap()
            .input
            .contains("Data!$D$2:$D$300")
    );

    for rows in [302, 300, 300, 300, 300, 300] {
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_dashboard(&workbook, rows);
    }
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_dashboard(&workbook, 300);
    assert_eq!(value(&workbook, 0, 1, 3), number(42.125));
    edit(&mut workbook, 0, "D2", "=21+22");
    assert_dashboard(&workbook, 300);
    assert!(
        workbook
            .cell(SheetId(0), CellRef::new(1, 3))
            .unwrap()
            .is_formula
    );
    edit(&mut workbook, 0, "D2", "44");
    assert_dashboard(&workbook, 300);
    assert!(
        !workbook
            .cell(SheetId(0), CellRef::new(1, 3))
            .unwrap()
            .is_formula
    );
}

#[test]
fn criteria_index_public_edits_match_current_data_oracle() {
    exercise_edits(false);
}

#[test]
fn criteria_index_public_edits_preserve_text_value_errors() {
    exercise_edits(true);
}

fn assert_fresh_dashboard(workbook: &Workbook) {
    assert_fresh_dashboard_rows(workbook, 300);
}

fn assert_fresh_dashboard_rows(workbook: &Workbook, data_rows: u32) {
    let bytes = workbook.save().unwrap();
    let mut fresh = Workbook::open(&bytes).unwrap();
    fresh.recalculate_all(CalculationOptions::default());
    for row in 0..=8 {
        for col in 0..=25 {
            assert_eq!(
                value(workbook, 1, row, col),
                value(&fresh, 1, row, col),
                "fresh-open dashboard {row}, {col}"
            );
        }
    }
    if !(1..=data_rows).any(|row| {
        (0..3).any(|col| matches!(value(workbook, 0, row, col), CellValue::Error { .. }))
    }) {
        assert_dashboard(workbook, data_rows);
        assert_dashboard(&fresh, data_rows);
    }
}

fn workbook_with_spill_anchor() -> Workbook {
    let mut model = workbook(false).model().clone();
    let data = &mut model.sheets[0];
    put(data, 0, 4, number(1.0));
    data.set_cell(
        CellRef::new(0, 1),
        Cell {
            formula: Some("_xlfn.SEQUENCE(E1)".into()),
            ..Cell::default()
        },
    );
    data.set_array_formula(CellRef::new(0, 1), CellRange::parse_a1("B1:B1").unwrap());
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    assert_fresh_dashboard(&workbook);
    workbook
}

fn flip_edits(workbook: &mut Workbook, edits: &[(&str, &str)], batched: bool) {
    if batched {
        let edits: Vec<_> = edits
            .iter()
            .map(|(address, input)| CellInput {
                cell: CellRef::parse_a1(address).unwrap(),
                input: (*input).into(),
            })
            .collect();
        let result = workbook
            .edit_cells(SheetId(0), &edits, CalculationOptions::default())
            .unwrap();
        assert!(result.applied);
        assert!(result.limited_cells.is_empty());
        assert!(result.cycle_cells.is_empty());
        assert_fresh_dashboard(workbook);
    } else {
        for (address, input) in edits {
            edit(workbook, 0, address, input);
            assert_fresh_dashboard(workbook);
        }
    }
}

fn exercise_unusable_range_flips(batched: bool) {
    let mut workbook = workbook_with_spill_anchor();
    flip_edits(
        &mut workbook,
        &[
            ("D2", "n/a"),
            ("A3", "=1/0"),
            ("D4", "=1+1"),
            ("B2", ""),
            ("B3", ""),
            ("E1", "3"),
        ],
        batched,
    );
    assert_eq!(value(&workbook, 0, 1, 3), text("n/a"));
    assert_eq!(
        value(&workbook, 0, 2, 0),
        CellValue::Error {
            value: ErrorValue::Div0
        }
    );
    assert_eq!(value(&workbook, 0, 3, 3), number(2.0));
    assert_eq!(value(&workbook, 0, 1, 1), number(2.0));
    assert_eq!(value(&workbook, 0, 2, 1), number(3.0));
    assert_eq!(
        workbook.model().sheets[0].array_formula(CellRef::new(0, 1)),
        Some(CellRange::parse_a1("B1:B3").unwrap())
    );

    flip_edits(
        &mut workbook,
        &[("A3", "a2"), ("D4", "0.375"), ("E1", "1"), ("D2", "0.125")],
        batched,
    );
    assert_eq!(
        workbook.model().sheets[0].array_formula(CellRef::new(0, 1)),
        Some(CellRange::parse_a1("B1:B3").unwrap())
    );
    assert_eq!(value(&workbook, 0, 1, 1), number(1.0));
    assert_eq!(value(&workbook, 0, 2, 1), number(1.0));
    for (row, col) in [(1, 3), (2, 0), (3, 3)] {
        assert!(
            !workbook
                .cell(SheetId(0), CellRef::new(row, col))
                .unwrap()
                .is_formula
        );
    }

    if batched {
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_fresh_dashboard(&workbook);
        assert_eq!(value(&workbook, 0, 1, 1), number(2.0));
        workbook.redo(CalculationOptions::default()).unwrap();
        assert_fresh_dashboard(&workbook);
        assert_eq!(value(&workbook, 0, 1, 1), number(1.0));
    }
    flip_edits(
        &mut workbook,
        &[("B1", ""), ("B3", "b0"), ("B2", "b0")],
        batched,
    );
}

#[test]
fn criteria_index_unusable_ranges_match_fresh_open_after_separate_edits() {
    exercise_unusable_range_flips(false);
}

#[test]
fn criteria_index_unusable_ranges_match_fresh_open_after_batched_edits() {
    exercise_unusable_range_flips(true);
}

#[test]
fn criteria_index_individual_range_flips_match_fresh_open() {
    let mut workbook = workbook_with_spill_anchor();
    for (address, unusable, usable) in [
        ("D2", "n/a", "0.125"),
        ("A3", "=1/0", "a2"),
        ("D4", "=1+1", "0.375"),
    ] {
        flip_edits(&mut workbook, &[(address, unusable)], false);
        flip_edits(&mut workbook, &[(address, usable)], false);
    }
    flip_edits(&mut workbook, &[("B2", ""), ("B3", ""), ("E1", "3")], false);
    assert_eq!(value(&workbook, 0, 1, 1), number(2.0));
    assert_eq!(value(&workbook, 0, 2, 1), number(3.0));
    flip_edits(
        &mut workbook,
        &[("E1", "1"), ("B1", ""), ("B3", "b0"), ("B2", "b0")],
        false,
    );
    assert_eq!(value(&workbook, 1, 3, 1), number(1.625));
}

#[test]
fn criteria_index_budget_refusal_does_not_poison_the_next_pass() {
    let mut workbook = workbook_with_rows(false, 60_000);
    assert_fresh_dashboard_rows(&workbook, 60_000);
    let burners: Vec<_> = (0..10)
        .map(|row| CellInput {
            cell: CellRef::new(row, 4),
            input: format!(
                "=ROWS(_xlfn.SEQUENCE({}))",
                if row == 9 { 999_900 } else { 1_000_000 }
            ),
        })
        .collect();
    workbook
        .edit_cells(SheetId(0), &burners, CalculationOptions::default())
        .unwrap();
    assert_fresh_dashboard_rows(&workbook, 60_000);
    let result = workbook.recalculate_all(CalculationOptions::default());
    assert!(result.budget_exhausted);
    assert!(
        result
            .limited_cells
            .iter()
            .any(|address| address.sheet == SheetId(1))
    );
    for row in 0..10 {
        assert_eq!(
            value(&workbook, 0, row, 4),
            number(if row == 9 { 999_900.0 } else { 1_000_000.0 })
        );
    }
    assert_fresh_dashboard_rows(&workbook, 60_000);

    let mut recovery: Vec<_> = burners
        .iter()
        .map(|edit| CellInput {
            cell: edit.cell,
            input: String::new(),
        })
        .collect();
    recovery.push(CellInput {
        cell: CellRef::new(1, 3),
        input: "42.125".into(),
    });
    workbook
        .edit_cells(SheetId(0), &recovery, CalculationOptions::default())
        .unwrap();
    assert!(!workbook.last_calculation().budget_exhausted);
    assert!(workbook.last_calculation().limited_cells.is_empty());
    assert_fresh_dashboard_rows(&workbook, 60_000);
    edit(&mut workbook, 0, "A2", "a2");
    assert_fresh_dashboard_rows(&workbook, 60_000);
}
