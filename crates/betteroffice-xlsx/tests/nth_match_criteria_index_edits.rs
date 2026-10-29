use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, Op, Sheet, SheetId, Workbook,
    WorkbookModel,
};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn value(wb: &Workbook, sheet: u32, row: u32, col: u32) -> CellValue {
    wb.sheet(SheetId(sheet))
        .unwrap()
        .cell(CellRef::new(row, col))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn assert_matches(wb: &Workbook, rows: u32) {
    assert!(!wb.last_calculation().budget_exhausted);
    let key = value(wb, 1, 0, 0);
    let expected: Vec<_> = (1..=rows)
        .filter(|&row| value(wb, 0, row, 0) == key)
        .map(|row| number(row as f64))
        .collect();
    for n in 1..=265 {
        assert_eq!(
            value(wb, 1, n, 1),
            expected
                .get((n - 1) as usize)
                .cloned()
                .unwrap_or(CellValue::Text {
                    value: String::new()
                }),
            "match {n}"
        );
    }
}

#[test]
fn nth_match_public_edits_and_structural_undo_match_current_rows() {
    let mut data = Sheet::new("Data");
    for row in 1..=10_000 {
        data.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: number((row % 3) as f64),
                ..Cell::default()
            },
        );
    }
    let mut dashboard = Sheet::new("Dashboard");
    dashboard.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    for n in 1..=265 {
        let at = CellRef::new(n, 1);
        dashboard.set_cell(
            at,
            Cell {
                formula: Some(format!("IFERROR(SMALL(IF(Data!$A$2:$A$10001=$A$1,ROW(Data!$A$2:$A$10001)-ROW(Data!$A$2)+1),ROW({n}:{n})),\"\")")),
                ..Cell::default()
            },
        );
        dashboard.set_array_formula(at, CellRange::new(at, at));
    }
    let mut model = WorkbookModel::default();
    model.sheets.extend([data, dashboard]);
    let mut wb = Workbook::from_model(model).unwrap();
    wb.recalculate_all(CalculationOptions::default());
    assert_matches(&wb, 10_000);
    for (sheet, address, input) in [(1, "A1", "2"), (0, "A2", "2")] {
        wb.edit_cell(
            SheetId(sheet),
            CellRef::parse_a1(address).unwrap(),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
        assert_matches(&wb, 10_000);
    }
    wb.apply_ops(
        vec![Op::InsertRows {
            sheet: SheetId(0),
            at: 50,
            count: 2,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    assert_matches(&wb, 10_002);
    wb.apply_ops(
        vec![Op::DeleteRows {
            sheet: SheetId(0),
            at: 100,
            count: 3,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    assert_matches(&wb, 9_999);
    for rows in [10_002, 10_000, 10_000, 10_000] {
        wb.undo(CalculationOptions::default()).unwrap();
        assert_matches(&wb, rows);
    }
}
