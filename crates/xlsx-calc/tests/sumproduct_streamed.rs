use xlsx_calc::{EvalContext, evaluate, evaluate_array, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{
    Cell, CellProvider, CellRange, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook,
};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}
fn t(v: &str) -> CellValue {
    CellValue::Text { value: v.into() }
}
fn b(v: bool) -> CellValue {
    CellValue::Bool { value: v }
}
fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn put(sheet: &mut Sheet, a1: &str, value: CellValue) {
    sheet.set_cell(
        CellRef::parse_a1(a1).unwrap(),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn eval_on(wb: &Workbook, sheet: SheetId, src: &str) -> CellValue {
    let expr = parse_formula(src).unwrap();
    evaluate(&expr, &EvalContext::new(wb, sheet))
}

fn fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    for (i, v) in ["Apple", "apple", "BANANA", "apple", "5"]
        .iter()
        .enumerate()
    {
        put(&mut s, &format!("A{}", i + 1), t(v));
    }
    for (i, v) in [10.0, 20.0, 30.0, 40.0, 50.0, 60.0].iter().enumerate() {
        put(&mut s, &format!("B{}", i + 1), n(*v));
    }
    for (i, v) in [true, false, true, true, false, true].iter().enumerate() {
        put(&mut s, &format!("C{}", i + 1), b(*v));
    }
    for (i, v) in ["5", "6", "x", "7", "8", "9"].iter().enumerate() {
        put(&mut s, &format!("E{}", i + 1), t(v));
    }
    put(&mut s, "G1", e(ErrorValue::Div0));
    put(&mut s, "G2", n(5.0));
    put(&mut s, "G3", n(6.0));
    wb.sheets.push(s);
    wb
}

fn check(wb: &Workbook, cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval_on(wb, SheetId(0), src), *want, "{src}");
    }
}

fn eval_array_on(wb: &Workbook, sheet: SheetId, src: &str) -> CellValue {
    let expr = parse_formula(src).unwrap();
    evaluate_array(&expr, &EvalContext::new(wb, sheet)).into_scalar()
}

fn check_array(wb: &Workbook, cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval_array_on(wb, SheetId(0), src), *want, "array {src}");
    }
}

fn put_array_formula(sheet: &mut Sheet, a1: &str, formula: &str) {
    let at = CellRef::parse_a1(a1).unwrap();
    sheet.set_cell(
        at,
        Cell {
            value: CellValue::Empty,
            formula: Some(formula.to_string()),
            ..Cell::default()
        },
    );
    sheet.set_array_formula(at, CellRange::new(at, at));
}

fn check_tail_accumulation(array: bool, overflow: bool) {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    put(&mut sheet, "A1", n(if overflow { -1.7e308 } else { 1.0 }));
    put(&mut sheet, "C1", n(0.1));
    put(&mut sheet, "D1", n(0.2));
    wb.sheets.push(sheet);
    let args: &[&str] = if overflow {
        &["A1:A3+1e308"]
    } else {
        &["A1:A1000+0.1", "A1:A1000+C1:D1", "A1:A3*-0"]
    };
    for arg in args {
        let streamed = format!("SUMPRODUCT({arg})");
        let fallback = format!("SUMPRODUCT(IF(TRUE,{arg}))");
        let eval = if array { eval_array_on } else { eval_on };
        let CellValue::Number { value: expected } = eval(&wb, SheetId(0), &fallback) else {
            panic!("{fallback} is not a number");
        };
        assert!(expected.is_finite(), "{fallback}");
        let CellValue::Number { value: actual } = eval(&wb, SheetId(0), &streamed) else {
            panic!("{streamed} is not a number");
        };
        assert_eq!(
            actual.to_bits(),
            expected.to_bits(),
            "{streamed}: {actual} vs {expected}"
        );
    }
}

#[test]
fn streamed_sumproduct_tail_overflow_matches_fallback() {
    check_tail_accumulation(false, true);
}

#[test]
fn array_sumproduct_tail_overflow_matches_fallback() {
    check_tail_accumulation(true, true);
}

#[test]
fn streamed_sumproduct_tail_low_bits_match_fallback() {
    check_tail_accumulation(false, false);
}

#[test]
fn array_sumproduct_tail_low_bits_match_fallback() {
    check_tail_accumulation(true, false);
}

#[test]
fn streamed_sumproduct_matches_scalar_semantics() {
    let wb = fixture();
    check(
        &wb,
        &[
            ("SUMPRODUCT((A1:A6=\"apple\")*B1:B6)", n(70.0)),
            ("SUMPRODUCT((A1:A6=\"APPLE\")*B1:B6)", n(70.0)),
            ("SUMPRODUCT((A1:A6=5)*B1:B6)", n(0.0)),
            ("SUMPRODUCT((A1:A6=\"5\")*B1:B6)", n(50.0)),
            ("SUMPRODUCT((B1:B6>15)*B1:B6)", n(200.0)),
            ("SUMPRODUCT(--(B1:B6>15))", n(5.0)),
            ("SUMPRODUCT(--(A1:A6=\"apple\"),B1:B6)", n(70.0)),
            ("SUMPRODUCT((C1:C6)*1)", n(4.0)),
            ("SUMPRODUCT(C1:C6)", n(0.0)),
            ("SUMPRODUCT(--E1:E2)", n(11.0)),
            ("SUMPRODUCT(-E1:E2)", n(-11.0)),
            ("SUMPRODUCT((E1:E6)*1)", e(ErrorValue::Value)),
            ("SUMPRODUCT((A1:A6=\"\")*1)", n(1.0)),
            ("SUMPRODUCT(B1:B2%)", n(10.0 / 100.0 + 20.0 / 100.0)),
            ("SUMPRODUCT((A1:A6=\"apple\")*ABS($B$1))", n(30.0)),
            ("SUMPRODUCT(+(A1:A6=\"apple\")*1)", n(3.0)),
            ("SUMPRODUCT((A1:A3=\"apple\")*G1:G3)", e(ErrorValue::Div0)),
            ("SUMPRODUCT(G1:G3)", e(ErrorValue::Div0)),
        ],
    );
}

#[test]
fn streamed_sumproduct_broadcasts_like_the_array_path() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    for (i, v) in [1.0, 2.0, 3.0].iter().enumerate() {
        put(&mut s, &format!("A{}", i + 1), n(*v));
    }
    put(&mut s, "C1", n(10.0));
    put(&mut s, "D1", n(20.0));
    for (r, row) in [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]].iter().enumerate() {
        put(&mut s, &format!("E{}", r + 1), n(row[0]));
        put(&mut s, &format!("F{}", r + 1), n(row[1]));
    }
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT((A1:A3)*(C1:D1))", n(180.0)),
            ("SUMPRODUCT((A1:A3)*(E1:F3))", n(50.0)),
            ("SUMPRODUCT((A1:A3=E1:F1)*1)", n(2.0)),
            ("SUMPRODUCT((A1:A3)*(A1:A3))", n(14.0)),
            ("SUMPRODUCT((A1:A2)*(A1:A3))", e(ErrorValue::NA)),
            ("SUMPRODUCT(A1:A2,A1:A3)", e(ErrorValue::Value)),
            ("SUMPRODUCT(A1:A3,C1:D1)", e(ErrorValue::Value)),
        ],
    );
}

fn two_sheet() -> Workbook {
    let mut wb = Workbook::default();
    let mut data = Sheet::new("Data");
    for (i, v) in ["x", "y", "x", "x"].iter().enumerate() {
        put(&mut data, &format!("A{}", i + 2), t(v));
    }
    for (i, v) in [1.0, 2.0, 3.0, 4.0].iter().enumerate() {
        put(&mut data, &format!("B{}", i + 2), n(*v));
    }
    let mut summary = Sheet::new("Summary");
    put(&mut summary, "A1", t("x"));
    wb.sheets.extend([data, summary]);
    wb
}

#[test]
fn streamed_sumproduct_reads_other_sheets_and_clips_tails() {
    let wb = two_sheet();
    let cases: &[(&str, CellValue)] = &[
        ("SUMPRODUCT((Data!A2:A5=$A$1)*(Data!B2:B5))", n(8.0)),
        ("SUMPRODUCT((Data!A:A=$A$1)*(Data!B:B))", n(8.0)),
        ("SUMPRODUCT((Data!A2:A100=$A$1)*(Data!B2:B100))", n(8.0)),
        ("SUMPRODUCT((Data!A2:A100=\"\")*1)", n(95.0)),
        ("SUMPRODUCT((Data!A2:A100=\"x\")*Data!B2:B100)", n(8.0)),
        ("SUMPRODUCT(1/(Data!A4:A100=\"x\"))", e(ErrorValue::Div0)),
    ];
    for (src, want) in cases {
        assert_eq!(eval_on(&wb, SheetId(1), src), *want, "{src}");
    }
}

#[test]
fn streamed_sumproduct_lifts_scalar_functions_over_ranges_via_fallback() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(10.0));
    put(&mut s, "A2", t("x"));
    put(&mut s, "A3", n(20.0));
    put(&mut s, "B1", n(1.0));
    put(&mut s, "B2", n(2.0));
    put(&mut s, "B3", n(3.0));
    put(&mut s, "D1", n(-10.0));
    put(&mut s, "D2", n(5.0));
    put(&mut s, "D3", n(-20.0));
    put(&mut s, "E1", n(1.0));
    put(&mut s, "E2", n(2.0));
    put(&mut s, "E3", n(1.0));
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT(ISNUMBER(A1:A3)*B1:B3)", n(4.0)),
            ("SUMPRODUCT((LEN(A1:A3)>0)*B1:B3)", n(6.0)),
            ("SUMPRODUCT(LEN(A1:A3))", n(5.0)),
            ("SUMPRODUCT(ABS(D1:D3))", n(35.0)),
            ("SUMPRODUCT(ABS(A1:A3))", e(ErrorValue::Value)),
            ("SUMPRODUCT(COUNTIF(E1:E3,E1:E3))", n(5.0)),
            ("SUMPRODUCT((A1:A3=\"x\")*ABS($B$1))", n(1.0)),
        ],
    );
}

#[test]
fn streamed_sumproduct_1x1_arguments_match_fallback_coercion() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", t("apple"));
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT(A1=\"apple\")", n(1.0)),
            ("SUMPRODUCT(A1:A1=\"apple\")", n(1.0)),
            ("SUMPRODUCT(\"5\",2)", n(10.0)),
            ("SUMPRODUCT(\"apple\")", e(ErrorValue::Value)),
            ("SUMPRODUCT(TRUE)", n(1.0)),
        ],
    );
}

#[test]
fn streamed_sumproduct_shape_checks_precede_tail_shortcut() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(1.0));
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT(A2:B10*A2:C10)", e(ErrorValue::NA)),
            ("SUMPRODUCT(A2:A10*A2:A9)", e(ErrorValue::NA)),
        ],
    );
}

#[test]
fn streamed_sumproduct_huge_declared_range_uses_tail_once() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(1.0));
    wb.sheets.push(s);
    let arg = "(A1:A1048576=\"\")*1";
    let formula = format!("SUMPRODUCT({})", vec![arg; 30].join(","));
    assert_eq!(eval_on(&wb, SheetId(0), &formula), n(0.0));
}

#[test]
fn streamed_sumproduct_empty_tail_and_error_precedence_match_fallback() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(5.0));
    put(&mut s, "A2", n(10.0));
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT((A1:A100=\"\")*1)", n(98.0)),
            ("SUMPRODUCT((A1:A100<>\"\")*1)", n(2.0)),
            ("SUMPRODUCT(A1:A100<>\"\")", n(0.0)),
            ("SUMPRODUCT(A1:A100=\"\")", n(0.0)),
            ("SUMPRODUCT(1/(A1:A100))", e(ErrorValue::Div0)),
        ],
    );
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", t("x"));
    wb.sheets.push(s);
    check(&wb, &[("SUMPRODUCT((A1:A100)*1)", e(ErrorValue::Value))]);
}

#[test]
fn streamed_sumproduct_whole_column_on_empty_sheet_matches_fallback() {
    let wb = {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb
    };
    check(
        &wb,
        &[
            ("SUMPRODUCT((A:A=\"\")*1)", n(1.0)),
            ("SUMPRODUCT((A:A<>\"\")*1)", n(0.0)),
            ("SUMPRODUCT((A1:A100=\"\")*1)", n(100.0)),
            ("SUMPRODUCT((A1:A100<>\"\")*1)", n(0.0)),
        ],
    );
}

#[test]
fn streamed_sumproduct_tail_error_precedes_later_argument_errors() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(1.0));
    put(&mut s, "B1", e(ErrorValue::NA));
    wb.sheets.push(s);
    check(
        &wb,
        &[
            ("SUMPRODUCT(1/A1:A10,B1:B10)", e(ErrorValue::Div0)),
            ("SUMPRODUCT(B1:B10,1/A1:A10)", e(ErrorValue::NA)),
        ],
    );
}

#[test]
fn streamed_sumproduct_one_row_tail_error_wins_on_recalc() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", e(ErrorValue::Div0));
    let at = CellRef::parse_a1("C1").unwrap();
    s.set_cell(
        at,
        Cell {
            value: n(99.0),
            formula: Some("SUMPRODUCT(A1:B1,A2:B2+VALUE(\"19/11/2021\"))".to_string()),
            ..Cell::default()
        },
    );
    wb.sheets.push(s);
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    assert_eq!(wb.value(SheetId(0), at), e(ErrorValue::Div0));
}

#[test]
fn streamed_sumproduct_reads_spill_interior_for_cycle_detection() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    let a1 = CellRef::parse_a1("A1").unwrap();
    s.set_cell(
        a1,
        Cell {
            value: CellValue::Empty,
            formula: Some("SUMPRODUCT(B2:B3*1)".to_string()),
            ..Cell::default()
        },
    );
    let b1 = CellRef::parse_a1("B1").unwrap();
    s.set_cell(
        b1,
        Cell {
            value: CellValue::Empty,
            formula: Some("A1+10".to_string()),
            ..Cell::default()
        },
    );
    s.set_array_formula(b1, CellRange::new(b1, CellRef::parse_a1("B2").unwrap()));
    wb.sheets.push(s);
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(result.cycle_cells, vec![(SheetId(0), a1), (SheetId(0), b1)]);
    assert_eq!(wb.value(SheetId(0), b1), t(""));
}

#[test]
fn array_sumproduct_matches_array_semantics() {
    let wb = fixture();
    check_array(
        &wb,
        &[
            ("SUMPRODUCT((A1:A6=\"apple\")*B1:B6)", n(70.0)),
            ("SUMPRODUCT((A1:A6=\"APPLE\")*B1:B6)", n(70.0)),
            ("SUMPRODUCT((A1:A6=5)*B1:B6)", n(0.0)),
            ("SUMPRODUCT((A1:A6=\"5\")*B1:B6)", n(50.0)),
            ("SUMPRODUCT((B1:B6>15)*B1:B6)", n(200.0)),
            ("SUMPRODUCT(--(B1:B6>15))", n(5.0)),
            ("SUMPRODUCT(--(A1:A6=\"apple\"),B1:B6)", n(70.0)),
            ("SUMPRODUCT((C1:C6)*1)", n(4.0)),
            ("SUMPRODUCT(C1:C6)", n(4.0)),
            ("SUMPRODUCT(--E1:E2)", n(11.0)),
            ("SUMPRODUCT(-E1:E2)", n(-11.0)),
            ("SUMPRODUCT((E1:E6)*1)", e(ErrorValue::Value)),
            ("SUMPRODUCT((A1:A6=\"\")*1)", n(1.0)),
            ("SUMPRODUCT(B1:B2%)", n(10.0 / 100.0 + 20.0 / 100.0)),
            ("SUMPRODUCT((A1:A6=\"apple\")*ABS($B$1))", n(30.0)),
            ("SUMPRODUCT(+(A1:A6=\"apple\")*1)", n(3.0)),
            ("SUMPRODUCT((A1:A3=\"apple\")*G1:G3)", e(ErrorValue::Div0)),
            ("SUMPRODUCT(G1:G3)", e(ErrorValue::Div0)),
            ("SUMPRODUCT((A1:A100=\"\")*1)", n(95.0)),
        ],
    );
}

#[test]
fn array_sumproduct_broadcasts_and_propagates_errors_by_position() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    for (i, v) in [1.0, 2.0, 3.0].iter().enumerate() {
        put(&mut s, &format!("A{}", i + 1), n(*v));
    }
    put(&mut s, "C1", n(10.0));
    put(&mut s, "D1", n(20.0));
    for (r, row) in [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]].iter().enumerate() {
        put(&mut s, &format!("E{}", r + 1), n(row[0]));
        put(&mut s, &format!("F{}", r + 1), n(row[1]));
    }
    put(&mut s, "H1", n(1.0));
    put(&mut s, "H2", n(1.0));
    put(&mut s, "H3", e(ErrorValue::Div0));
    put(&mut s, "I1", e(ErrorValue::NA));
    put(&mut s, "I2", n(1.0));
    put(&mut s, "I3", n(1.0));
    wb.sheets.push(s);
    check_array(
        &wb,
        &[
            ("SUMPRODUCT((A1:A3)*(C1:D1))", n(180.0)),
            ("SUMPRODUCT((A1:A3)*(E1:F3))", n(50.0)),
            ("SUMPRODUCT((A1:A3=E1:F1)*1)", n(2.0)),
            ("SUMPRODUCT((A1:A3)*(A1:A3))", n(14.0)),
            ("SUMPRODUCT((A1:A2)*(A1:A3))", e(ErrorValue::NA)),
            ("SUMPRODUCT(A1:A2,A1:A3)", e(ErrorValue::NA)),
            ("SUMPRODUCT(A1:A3,C1:D1)", n(180.0)),
            ("SUMPRODUCT(2,A1:A3)", n(12.0)),
            ("SUMPRODUCT(H1:H3,I1:I3)", e(ErrorValue::NA)),
        ],
    );
}

#[test]
fn array_sumproduct_lifted_scalar_functions_fall_back() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", n(10.0));
    put(&mut s, "A2", t("x"));
    put(&mut s, "A3", n(20.0));
    put(&mut s, "B1", n(1.0));
    put(&mut s, "B2", n(2.0));
    put(&mut s, "B3", n(3.0));
    put(&mut s, "D1", n(-10.0));
    put(&mut s, "D2", n(5.0));
    put(&mut s, "D3", n(-20.0));
    put(&mut s, "E1", n(1.0));
    put(&mut s, "E2", n(2.0));
    put(&mut s, "E3", n(1.0));
    wb.sheets.push(s);
    check_array(
        &wb,
        &[
            ("SUMPRODUCT(ISNUMBER(A1:A3)*B1:B3)", n(4.0)),
            ("SUMPRODUCT((LEN(A1:A3)>0)*B1:B3)", n(6.0)),
            ("SUMPRODUCT(LEN(A1:A3))", n(5.0)),
            ("SUMPRODUCT(ABS(D1:D3))", n(35.0)),
            ("SUMPRODUCT(ABS(A1:A3))", e(ErrorValue::Value)),
            ("SUMPRODUCT(COUNTIF(E1:E3,E1:E3))", n(5.0)),
            ("SUMPRODUCT((A1:A3=\"x\")*ABS($B$1))", n(1.0)),
        ],
    );
}

#[test]
fn array_formula_sumproduct_inside_if_matches() {
    let mut wb = Workbook::default();
    let mut data = Sheet::new("Data");
    for (i, v) in ["x", "y", "x", "x"].iter().enumerate() {
        put(&mut data, &format!("A{}", i + 2), t(v));
    }
    for (i, v) in ["Total", "Total", "Other", "Total"].iter().enumerate() {
        put(&mut data, &format!("G{}", i + 2), t(v));
    }
    for (i, v) in [1.0, 2.0, 3.0, 4.0].iter().enumerate() {
        put(&mut data, &format!("J{}", i + 2), n(*v));
    }
    let mut summary = Sheet::new("Summary");
    put(&mut summary, "A2", t("Total"));
    put(&mut summary, "A5", t("x"));
    put(&mut summary, "A4", t("x"));
    put_array_formula(
        &mut summary,
        "C1",
        "IF($A$2=\"Total\",SUMPRODUCT((Data!$A$2:$A$5=$A5)*Data!$J$2:$J$5),SUMPRODUCT((Data!$A$2:$A$5=$A5)*(Data!$G$2:$G$5=$A$2)*Data!$J$2:$J$5))",
    );
    put_array_formula(
        &mut summary,
        "C2",
        "IF($A$4=\"2009/10\",SUM(Data!$J$2:$J$5),SUMPRODUCT((Data!$A$2:$A$5=$A$4)*Data!$J$2:$J$5))",
    );
    wb.sheets.extend([data, summary]);
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    assert_eq!(
        wb.value(SheetId(1), CellRef::parse_a1("C1").unwrap()),
        n(8.0)
    );
    assert_eq!(
        wb.value(SheetId(1), CellRef::parse_a1("C2").unwrap()),
        n(8.0)
    );
    put(wb.sheet_mut(SheetId(1)).unwrap(), "A2", t("Other"));
    put(wb.sheet_mut(SheetId(1)).unwrap(), "A4", t("2009/10"));
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    assert_eq!(
        wb.value(SheetId(1), CellRef::parse_a1("C1").unwrap()),
        n(3.0)
    );
    assert_eq!(
        wb.value(SheetId(1), CellRef::parse_a1("C2").unwrap()),
        n(10.0)
    );
}

#[test]
fn many_array_sumproducts_fit_the_recalc_budget() {
    let mut wb = Workbook::default();
    let mut data = Sheet::new("Data");
    for i in 0..50_000u32 {
        data.set_cell(
            CellRef::new(i, 0),
            Cell {
                value: t(if i % 2 == 0 { "x" } else { "z" }),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 1),
            Cell {
                value: n(if i % 3 == 0 { 7.0 } else { 8.0 }),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 2),
            Cell {
                value: t("y"),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 3),
            Cell {
                value: n(2.0),
                ..Cell::default()
            },
        );
    }
    let mut summary = Sheet::new("Summary");
    put(&mut summary, "A1", t("x"));
    put(&mut summary, "B1", n(7.0));
    let formula = "IF($A$1=\"x\",SUMPRODUCT((Data!$A$1:$A$100000=$A$1)*(Data!$B$1:$B$100000=$B$1)*(Data!$C$1:$C$100000=\"y\")*(Data!$D$1:$D$100000)),SUMPRODUCT((Data!$A$1:$A$100000=$A$1)*(Data!$D$1:$D$100000)))";
    for i in 0..40u32 {
        let at = CellRef::new(i + 1, 2);
        summary.set_cell(
            at,
            Cell {
                value: CellValue::Empty,
                formula: Some(formula.to_string()),
                ..Cell::default()
            },
        );
        summary.set_array_formula(at, CellRange::new(at, at));
    }
    wb.sheets.extend([data, summary]);
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    for i in 0..40u32 {
        assert_eq!(
            wb.value(SheetId(1), CellRef::new(i + 1, 2)),
            n(16_668.0),
            "array formula row {}",
            i + 2
        );
    }
}

#[test]
fn many_streamed_sumproducts_fit_the_recalc_budget() {
    let mut wb = Workbook::default();
    let mut data = Sheet::new("Data");
    for i in 0..50_000u32 {
        let row = CellRef::new(i, 0);
        let key = if i % 2 == 0 { "x" } else { "z" };
        data.set_cell(
            row,
            Cell {
                value: t(key),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 1),
            Cell {
                value: n(if i % 3 == 0 { 7.0 } else { 8.0 }),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 2),
            Cell {
                value: t("y"),
                ..Cell::default()
            },
        );
        data.set_cell(
            CellRef::new(i, 3),
            Cell {
                value: n(2.0),
                ..Cell::default()
            },
        );
    }
    let mut summary = Sheet::new("Summary");
    put(&mut summary, "A1", t("x"));
    put(&mut summary, "B1", n(7.0));
    let formula = "SUMPRODUCT((Data!$A$1:$A$100000=$A$1)*(Data!$B$1:$B$100000=$B$1)*(Data!$C$1:$C$100000=\"y\")*(Data!$D$1:$D$100000))";
    for i in 0..40u32 {
        summary.set_cell(
            CellRef::new(i + 1, 2),
            Cell {
                value: CellValue::Empty,
                formula: Some(formula.to_string()),
                ..Cell::default()
            },
        );
    }
    wb.sheets.extend([data, summary]);
    let (_graph, result) = rebuild_and_recalc_all(&mut wb, None);
    assert!(
        result.limited_cells.is_empty(),
        "{:?}",
        result.limited_cells
    );
    for i in 0..40u32 {
        assert_eq!(
            wb.value(SheetId(1), CellRef::new(i + 1, 2)),
            n(16_668.0),
            "formula row {}",
            i + 2
        );
    }
}
