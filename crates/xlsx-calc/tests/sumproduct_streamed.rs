use xlsx_calc::{EvalContext, evaluate, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

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
    for (i, v) in ["Apple", "apple", "BANANA", "apple", "5"].iter().enumerate() {
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
            (
                "SUMPRODUCT(B1:B2%)",
                n(10.0 / 100.0 + 20.0 / 100.0),
            ),
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
            ("SUMPRODUCT((A1:A3)*(C1:D1))", n(210.0)),
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
        (
            "SUMPRODUCT(1/(Data!A4:A100=\"x\"))",
            e(ErrorValue::Div0),
        ),
    ];
    for (src, want) in cases {
        assert_eq!(eval_on(&wb, SheetId(1), src), *want, "{src}");
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
    assert!(result.limited_cells.is_empty(), "{:?}", result.limited_cells);
    for i in 0..40u32 {
        assert_eq!(
            wb.value(SheetId(1), CellRef::new(i + 1, 2)),
            n(16_668.0),
            "formula row {}",
            i + 2
        );
    }
}
