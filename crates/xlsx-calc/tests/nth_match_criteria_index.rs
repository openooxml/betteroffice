use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellProvider, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

#[test]
fn nth_match_cse_scale_fits_default_budget() {
    let mut data = Sheet::new("Data");
    let mut expected = Vec::new();
    for row in 0..10_000 {
        let key = if row % 20 == 0 { 7 } else { 100 + row % 300 };
        if key == 7 {
            expected.push(row);
        }
        data.set_cell(
            CellRef::new(row + 5, 3),
            Cell {
                value: number(key as f64),
                ..Cell::default()
            },
        );
        for col in 4..9 {
            data.set_cell(
                CellRef::new(row + 5, col),
                Cell {
                    value: number((row * 10 + col) as f64),
                    ..Cell::default()
                },
            );
        }
    }
    let mut dashboard = Sheet::new("Dashboard");
    dashboard.set_cell(
        CellRef::new(2, 1),
        Cell {
            value: number(7.0),
            ..Cell::default()
        },
    );
    for n in 1..=265 {
        for col in 0..5 {
            let letter = char::from(b'E' + col as u8);
            let lookup = format!(
                "INDEX('Data'!${letter}$6:${letter}$10005,SMALL(IF($B$3='Data'!$D$6:$D$10005,ROW('Data'!$D$6:$D$10005)-ROW('Data'!$D$6)+1),ROW({n}:{n})))"
            );
            let at = CellRef::new(n + 4, col);
            dashboard.set_cell(
                at,
                Cell {
                    formula: Some(format!("IF(ISERROR({lookup}),\"\",{lookup})")),
                    ..Cell::default()
                },
            );
            dashboard.set_array_formula(at, CellRange::new(at, at));
        }
    }
    for col in 0..5 {
        let letter = char::from(b'A' + col as u8);
        for (row, source) in [
            (271, format!("COUNTIF({letter}6:{letter}270,\">0\")")),
            (
                272,
                format!("SUMIF({letter}6:{letter}270,\">0\",{letter}6:{letter}270)"),
            ),
        ] {
            dashboard.set_cell(
                CellRef::new(row, col),
                Cell {
                    formula: Some(source),
                    ..Cell::default()
                },
            );
        }
    }
    let mut wb = Workbook::default();
    wb.sheets.extend([data, dashboard]);
    let (_, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(!result.budget_exhausted);
    assert!(result.limited_cells.is_empty());
    for col in 0..5 {
        let mut total = 0.0;
        for n in 1..=265 {
            let value = (expected[(n - 1) as usize] * 10 + col + 4) as f64;
            total += value;
            assert_eq!(
                wb.value(SheetId(1), CellRef::new(n + 4, col)),
                number(value)
            );
        }
        assert_eq!(wb.value(SheetId(1), CellRef::new(271, col)), number(265.0));
        assert_eq!(wb.value(SheetId(1), CellRef::new(272, col)), number(total));
    }
}
