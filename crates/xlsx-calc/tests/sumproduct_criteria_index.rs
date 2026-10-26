use xlsx_calc::{EvalContext, evaluate, parse_formula, rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

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

fn formula(sheet: &mut Sheet, row: u32, col: u32, source: &str) {
    sheet.set_cell(
        CellRef::new(row, col),
        Cell {
            formula: Some(source.into()),
            ..Cell::default()
        },
    );
}

fn eval(wb: &Workbook, source: &str) -> CellValue {
    evaluate(
        &parse_formula(source).unwrap(),
        &EvalContext::new(wb, SheetId(0)),
    )
}

fn equivalent(wb: &Workbook, argument: &str) {
    let actual = eval(wb, &format!("SUMPRODUCT({argument})"));
    let expected = eval(wb, &format!("SUMPRODUCT(({argument})+0)"));
    match (&actual, &expected) {
        (CellValue::Number { value: a }, CellValue::Number { value: b }) => {
            assert_eq!(a.to_bits(), b.to_bits(), "{argument}");
        }
        _ => assert_eq!(actual, expected, "{argument}"),
    }
}

fn scale() -> Workbook {
    let mut data = Sheet::new("Data");
    for row in 0..20_000 {
        put(&mut data, row, 0, text(&format!("a{}", row % 100)));
        put(&mut data, row, 1, text(&format!("b{}", row % 37)));
        put(&mut data, row, 2, number((row % 13) as f64));
        put(&mut data, row, 3, number((row % 17) as f64 / 8.0));
    }
    let mut dashboard = Sheet::new("Dashboard");
    for row in 0..1_500 {
        formula(
            &mut dashboard,
            row,
            0,
            &format!(
                "IF(TRUE,SUMPRODUCT((Data!A1:A20000=\"a{}\")*(Data!B1:B20000=\"b{}\")*(Data!C1:C20000={})*Data!D1:D20000),0)",
                row % 100,
                row % 37,
                row % 13
            ),
        );
    }
    let mut wb = Workbook::default();
    wb.sheets.extend([data, dashboard]);
    wb
}

#[test]
fn criteria_index_scale_matches_rust_loop_without_budget_refusal() {
    let mut wb = scale();
    let (_, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(!result.budget_exhausted);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    for key in 0..1_500 {
        let mut expected = 0.0;
        for row in 0..20_000 {
            if row % 100 == key % 100 && row % 37 == key % 37 && row % 13 == key % 13 {
                expected += (row % 17) as f64 / 8.0;
            }
        }
        assert_eq!(
            wb.value(SheetId(1), CellRef::new(key, 0)),
            number(expected),
            "dashboard row {key}"
        );
    }
}

#[test]
fn criteria_index_cse_scale_matches_rust_loop_without_budget_refusal() {
    let mut wb = scale();
    for row in 0..1_500 {
        let at = CellRef::new(row, 0);
        let product = format!(
            "SUMPRODUCT((Data!A1:A20000=\"a{}\")*(Data!B1:B20000=\"b{}\")*(Data!C1:C20000={})*Data!D1:D20000)",
            row % 100,
            row % 37,
            row % 13
        );
        formula(
            &mut wb.sheets[1],
            row,
            0,
            &format!("IF(MOD({},2)=0,{product},{product})", row),
        );
        wb.sheets[1].set_array_formula(at, xlsx_model::CellRange::new(at, at));
    }
    let (_, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(!result.budget_exhausted);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    for key in 0..1_500 {
        let mut expected = 0.0;
        for row in 0..20_000 {
            if row % 100 == key % 100 && row % 37 == key % 37 && row % 13 == key % 13 {
                expected += (row % 17) as f64 / 8.0;
            }
        }
        assert_eq!(
            wb.value(SheetId(1), CellRef::new(key, 0)),
            number(expected),
            "dashboard row {key}"
        );
    }
}

fn mixed() -> Workbook {
    let mut sheet = Sheet::new("Data");
    for (row, value) in [
        number(5.0),
        text("5"),
        CellValue::Bool { value: false },
        CellValue::Bool { value: true },
        text("Apple"),
        text("APPLE"),
        number(-0.0),
        number(0.0),
        text(""),
        CellValue::Empty,
    ]
    .into_iter()
    .enumerate()
    {
        put(&mut sheet, row as u32, 0, value);
        put(&mut sheet, row as u32, 1, number((row + 1) as f64 / 8.0));
    }
    let mut keys = Sheet::new("Keys");
    put(&mut keys, 0, 0, text("apple"));
    let mut wb = Workbook::default();
    wb.sheets.extend([sheet, keys]);
    wb
}

#[test]
fn criteria_index_mixed_keys_match_existing_equality() {
    let wb = mixed();
    for key in [
        "5",
        "\"5\"",
        "FALSE",
        "TRUE",
        "\"aPpLe\"",
        "0",
        "-0",
        "\"\"",
        "Keys!A1",
        "Keys!A2",
        "1/0",
        "2+3",
        "LOWER(Keys!A1)",
    ] {
        equivalent(&wb, &format!("(A1:A10={key})*B1:B10"));
        equivalent(&wb, &format!("--({key}=A1:A10)*(B1:B10)"));
    }
    assert_eq!(
        eval(&wb, "IF(TRUE,SUMPRODUCT((A1:A10=Keys!A1)*B1:B10),0)"),
        number(11.0 / 8.0)
    );
}

#[test]
fn criteria_index_refusals_preserve_existing_results() {
    let mut wb = mixed();
    put(
        &mut wb.sheets[0],
        0,
        0,
        CellValue::Error {
            value: ErrorValue::Div0,
        },
    );
    equivalent(&wb, "(A1:A10=5)*B1:B10");
    put(&mut wb.sheets[0], 0, 0, number(5.0));
    put(&mut wb.sheets[0], 0, 1, text("bad"));
    equivalent(&wb, "(A1:A10=5)*B1:B10");
    formula(&mut wb.sheets[0], 0, 1, "2+3");
    rebuild_and_recalc_all(&mut wb, None);
    equivalent(&wb, "(A1:A10=5)*B1:B10");
    equivalent(&wb, "(A1:A9=5)*B1:B10");
    equivalent(&wb, "(A1:A10=A1:A10)*B1:B10");
}

#[test]
fn criteria_index_incremental_edit_uses_fresh_cache() {
    let mut wb = mixed();
    formula(
        &mut wb.sheets[1],
        0,
        1,
        "SUMPRODUCT((Data!A1:A10=5)*Data!B1:B10)",
    );
    let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(wb.value(SheetId(1), CellRef::new(0, 1)), number(0.125));
    put(&mut wb.sheets[0], 0, 1, number(42.0));
    let result = recalc_after(
        &mut wb,
        &mut graph,
        &[(SheetId(0), CellRef::new(0, 1))],
        None,
    );
    assert!(result.limited_cells.is_empty());
    assert!(!result.budget_exhausted);
    assert_eq!(wb.value(SheetId(1), CellRef::new(0, 1)), number(42.0));
    put(&mut wb.sheets[0], 1, 0, number(5.0));
    recalc_after(
        &mut wb,
        &mut graph,
        &[(SheetId(0), CellRef::new(1, 0))],
        None,
    );
    assert_eq!(wb.value(SheetId(1), CellRef::new(0, 1)), number(42.25));
}

#[test]
fn criteria_index_horizontal_products_preserve_summation_bits() {
    let mut sheet = Sheet::new("Data");
    for (col, value) in [1e16, 0.1, -1e16, 0.2, -0.0, 0.3].into_iter().enumerate() {
        put(&mut sheet, 0, col as u32, text("x"));
        put(&mut sheet, 1, col as u32, number(value));
        put(&mut sheet, 2, col as u32, number(5.0));
    }
    let mut wb = Workbook::default();
    wb.sheets.push(sheet);
    for argument in [
        "(A1:F1=\"X\")*A2:F2",
        "(A1:F1=\"x\")*(A3:F3=5)*A2:F2",
        "A2:F2*((A1:F1=\"x\")*(A3:F3=5))",
        "--(A1:F1=\"x\")*--(5=A3:F3)*A2:F2",
    ] {
        equivalent(&wb, argument);
    }
}

fn spill_into_indexed_range(e1: f64) -> Workbook {
    let mut data = Sheet::new("Data");
    for (row, value) in [(1, 10.0), (2, 20.0), (3, 30.0)] {
        put(&mut data, row, 2, number(value));
    }
    put(&mut data, 0, 4, number(e1));
    formula(&mut data, 0, 0, "SUMPRODUCT((B2:B4=2)*C2:C4)");
    formula(&mut data, 0, 3, "SUMPRODUCT(((B2:B4=2)*C2:C4)+0)");
    formula(&mut data, 0, 1, "_xlfn.SEQUENCE(E1)");
    data.set_array_formula(
        CellRef::new(0, 1),
        xlsx_model::CellRange::new(CellRef::new(0, 1), CellRef::new(0, 1)),
    );
    let mut wb = Workbook::default();
    wb.sheets.push(data);
    wb
}

#[test]
fn criteria_index_sees_values_a_later_spill_round_writes() {
    let mut wb = spill_into_indexed_range(3.0);
    rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(wb.value(SheetId(0), CellRef::new(2, 1)), number(3.0));
    assert_eq!(wb.value(SheetId(0), CellRef::new(0, 3)), number(10.0));
    assert_eq!(wb.value(SheetId(0), CellRef::new(0, 0)), number(10.0));

    let mut wb = spill_into_indexed_range(1.0);
    let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(wb.value(SheetId(0), CellRef::new(0, 0)), number(0.0));
    put(&mut wb.sheets[0], 0, 4, number(3.0));
    recalc_after(
        &mut wb,
        &mut graph,
        &[(SheetId(0), CellRef::new(0, 4))],
        None,
    );
    assert_eq!(wb.value(SheetId(0), CellRef::new(0, 3)), number(10.0));
    assert_eq!(wb.value(SheetId(0), CellRef::new(0, 0)), number(10.0));
}

fn same_round_spill(unsupported: bool, indexed: bool) -> Workbook {
    let mut wb = spill_into_indexed_range(3.0);
    let sheet = &mut wb.sheets[0];
    let product = if indexed {
        "SUMPRODUCT((B2:B4=2)*C2:C4)"
    } else {
        "SUMPRODUCT(((B2:B4=2)*C2:C4)+0)"
    };
    formula(sheet, 0, 0, product);
    if unsupported {
        formula(
            sheet,
            0,
            3,
            &format!("IF({product}=10,WEBSERVICE(\"x\"),99)"),
        );
        sheet.cell_mut(CellRef::new(0, 3)).unwrap().value = number(42.0);
    } else {
        for (row, col, source) in [
            (0, 6, format!("_xlfn.SEQUENCE(IF({product}=10,1,2))")),
            (1, 5, "_xlfn.SEQUENCE(1,2)".into()),
        ] {
            formula(sheet, row, col, &source);
            let at = CellRef::new(row, col);
            sheet.set_array_formula(at, xlsx_model::CellRange::new(at, at));
        }
    }
    wb
}

#[test]
fn criteria_index_invalidates_before_same_round_spill_blocking() {
    let mut actual = same_round_spill(false, true);
    let mut expected = same_round_spill(false, false);
    rebuild_and_recalc_all(&mut actual, None);
    rebuild_and_recalc_all(&mut expected, None);
    for row in 0..4 {
        for col in 0..7 {
            let at = CellRef::new(row, col);
            assert_eq!(actual.value(SheetId(0), at), expected.value(SheetId(0), at));
            assert_eq!(
                actual.sheets[0].array_formula(at),
                expected.sheets[0].array_formula(at)
            );
        }
    }
    assert_eq!(actual.value(SheetId(0), CellRef::new(1, 5)), number(1.0));
    assert_eq!(actual.value(SheetId(0), CellRef::new(1, 6)), number(2.0));
    assert_eq!(
        actual.sheets[0].array_formula(CellRef::new(1, 5)),
        Some(xlsx_model::CellRange::new(
            CellRef::new(1, 5),
            CellRef::new(1, 6)
        ))
    );
}

#[test]
fn criteria_index_invalidates_before_unsupported_branch() {
    let mut actual = same_round_spill(true, true);
    let mut expected = same_round_spill(true, false);
    rebuild_and_recalc_all(&mut actual, None);
    rebuild_and_recalc_all(&mut expected, None);
    for at in [CellRef::new(0, 0), CellRef::new(0, 3)] {
        assert_eq!(actual.value(SheetId(0), at), expected.value(SheetId(0), at));
    }
    assert_eq!(actual.value(SheetId(0), CellRef::new(0, 3)), number(42.0));
}

#[test]
fn criteria_index_whole_column_style_tails_match_existing_path() {
    let wb = mixed();
    for argument in [
        "(A2:A1000000=1)*B2:B1000000",
        "(A2:A1000000=\"\")*B2:B1000000",
        "(A20:A1000000=1)*B20:B1000000",
        "(A20:A1000000=\"\")*B20:B1000000",
    ] {
        equivalent(&wb, argument);
    }
}
