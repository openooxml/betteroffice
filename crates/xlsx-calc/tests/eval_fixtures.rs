//! table-driven end-to-end tests: parse a formula, evaluate it against a small
//! in-memory workbook (which implements `CellProvider`), assert the value.

use xlsx_calc::{EvalContext, evaluate, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

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

/// two-sheet fixture: Sheet1 has A1..A3 = 10/20/30, B1 = "hi", B2 = TRUE,
/// D1 = "5" (C1 empty); Data has A1 = 100.
fn fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut s1 = Sheet::new("Sheet1");
    let put = |s: &mut Sheet, a1: &str, v: CellValue| {
        s.set_cell(
            CellRef::parse_a1(a1).unwrap(),
            Cell {
                value: v,
                ..Cell::default()
            },
        );
    };
    put(&mut s1, "A1", n(10.0));
    put(&mut s1, "A2", n(20.0));
    put(&mut s1, "A3", n(30.0));
    put(&mut s1, "B1", t("hi"));
    put(&mut s1, "B2", b(true));
    put(&mut s1, "D1", t("5"));
    wb.sheets.push(s1);

    let mut data = Sheet::new("Data");
    put(&mut data, "A1", n(100.0));
    wb.sheets.push(data);
    wb
}

fn eval(src: &str) -> CellValue {
    let wb = fixture();
    let expr = parse_formula(src).expect("parse");
    let ctx = EvalContext::new(&wb, SheetId(0));
    evaluate(&expr, &ctx)
}

#[test]
fn arithmetic_and_precedence() {
    let cases: &[(&str, CellValue)] = &[
        ("1+2*3", n(7.0)),
        ("(1+2)*3", n(9.0)),
        ("2^3^2", n(64.0)), // excel: left-associative
        ("-2^2", n(4.0)),   // excel: (-2)^2
        ("10/4", n(2.5)),
        ("10/0", e(ErrorValue::Div0)),
        ("50%", n(0.5)),
        ("-50%", n(-0.5)),
        ("2+2=4", b(true)),
        ("2<>3", b(true)),
        ("3<=3", b(true)),
    ];
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

#[test]
fn coercion_rules() {
    let cases: &[(&str, CellValue)] = &[
        ("TRUE+1", n(2.0)),
        ("FALSE*5", n(0.0)),
        ("C1+5", n(5.0)),
        ("\"5\"+2", n(7.0)),
        ("\"x\"+2", e(ErrorValue::Value)),
        ("1&2", t("12")),
        ("\"a\"&TRUE", t("aTRUE")),
    ];
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

#[test]
fn error_propagation() {
    assert_eq!(eval("#REF! + 1"), e(ErrorValue::Ref));
    assert_eq!(eval("1 + #DIV/0!"), e(ErrorValue::Div0));
    assert_eq!(eval("#N/A + #VALUE!"), e(ErrorValue::NA));
}

#[test]
fn references_and_sheets() {
    let cases: &[(&str, CellValue)] = &[
        ("A1", n(10.0)),
        ("A1+A2", n(30.0)),
        ("B1", t("hi")),
        ("Data!A1", n(100.0)),
        ("Data!A1*2", n(200.0)),
        ("Nope!A1", e(ErrorValue::Ref)),
        ("Z99", CellValue::Empty),
        ("Z99+1", n(1.0)),
    ];
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

#[test]
fn functions() {
    let cases: &[(&str, CellValue)] = &[
        ("SUM(A1:A3)", n(60.0)),
        ("SUM(A1:A3, 100)", n(160.0)),
        ("SUM(B1:B2)", n(0.0)),
        ("SUM(TRUE, 1)", n(2.0)),
        ("AVERAGE(A1:A3)", n(20.0)),
        ("AVERAGE(B1)", e(ErrorValue::Div0)),
        ("COUNT(A1:B2)", n(2.0)),
        ("COUNTA(A1:B2)", n(4.0)),
        ("MIN(A1:A3)", n(10.0)),
        ("MAX(A1:A3)", n(30.0)),
        ("IF(A1>5, \"big\", \"small\")", t("big")),
        ("IF(A1>50, 1)", b(false)),
        ("AND(TRUE, A1>5)", b(true)),
        ("OR(FALSE, A1>50)", b(false)),
        ("NOT(A1>50)", b(true)),
        ("ABS(-7)", n(7.0)),
        ("ROUND(2.345, 2)", n(2.35)),
        ("ROUND(2.5, 0)", n(3.0)),
        ("LEN(B1)", n(2.0)),
        ("CONCATENATE(\"a\", 1, TRUE)", t("a1TRUE")),
        ("CONCAT(A1:A2)", t("1020")),
        ("TRIM(\"  a   b  \")", t("a b")),
        ("UPPER(B1)", t("HI")),
        ("LOWER(\"HeLLo\")", t("hello")),
        ("D1+1", n(6.0)),
        ("NOSUCHFUNC(1)", e(ErrorValue::Name)),
    ];
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

#[test]
fn malformed_inputs_error_not_panic() {
    for src in ["", "1+", "SUM(", "(1", "1 2", ")", "&1"] {
        assert!(parse_formula(src).is_err(), "should reject {src:?}");
    }
}
