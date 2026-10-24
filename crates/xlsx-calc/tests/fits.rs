//! the least-squares fit family (RSQ, SLOPE, INTERCEPT, CORREL) against the
//! Microsoft worked example and the template real workbooks use it in.

use xlsx_calc::{EvalContext, evaluate, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}
fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn put(s: &mut Sheet, a1: &str, value: CellValue) {
    s.set_cell(
        CellRef::parse_a1(a1).unwrap(),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

/// twelve months of energy against throughput in E46:P47, the same template
/// left empty in E56:P57, and the RSQ example in A2:B8.
fn fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut fit = Sheet::new("Fit");
    let energy = [
        470000.0, 470000.0, 470000.0, 470000.0, 470000.0, 470000.0, 470000.0, 470000.0, 470000.0,
        470000.0, 470000.0, 492946.8,
    ];
    let throughput = [
        11400.0, 11400.0, 11400.0, 11400.0, 11400.0, 11300.0, 11400.0, 11400.0, 11400.0, 11400.0,
        11400.0, 11823.86,
    ];
    for (i, column) in ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
        .iter()
        .enumerate()
    {
        put(&mut fit, &format!("{column}46"), n(energy[i]));
        put(&mut fit, &format!("{column}47"), n(throughput[i]));
    }
    for (i, (y, x)) in [2.0, 3.0, 9.0, 1.0, 8.0, 7.0, 5.0]
        .iter()
        .zip([6.0, 5.0, 11.0, 7.0, 5.0, 4.0, 4.0])
        .enumerate()
    {
        put(&mut fit, &format!("A{}", i + 2), n(*y));
        put(&mut fit, &format!("B{}", i + 2), n(x));
    }
    wb.sheets.push(fit);
    wb
}

fn eval(src: &str) -> CellValue {
    let wb = fixture();
    let expr = parse_formula(src).expect("parse");
    let ctx = EvalContext::new(&wb, SheetId(0));
    evaluate(&expr, &ctx)
}

fn near(got: CellValue, want: f64, tolerance: f64, src: &str) {
    match got {
        CellValue::Number { value } => assert!(
            (value - want).abs() <= tolerance,
            "formula {src:?}: {value} != {want} (tolerance {tolerance})"
        ),
        other => panic!("formula {src:?}: expected a number, got {other:?}"),
    }
}

fn check(cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

/// an energy-correction template fits monthly energy against throughput with
/// `RSQ(E46:P46,E47:P47)`; left empty, excel shows `#DIV/0!` for the whole fit.
#[test]
fn least_squares_fits_on_a_monthly_template() {
    near(
        eval("RSQ(E46:P46,E47:P47)"),
        0.9497509010874818,
        1e-12,
        "RSQ",
    );
    near(
        eval("SLOPE(E46:P46,E47:P47)"),
        50.33767921364506,
        1e-9,
        "SLOPE",
    );
    near(
        eval("INTERCEPT(E46:P46,E47:P47)"),
        -103295.83976806462,
        1e-6,
        "INTERCEPT",
    );
    check(&[
        ("SLOPE(E56:P56,E57:P57)", e(ErrorValue::Div0)),
        ("INTERCEPT(E56:P56,E57:P57)", e(ErrorValue::Div0)),
        ("RSQ(E56:P56,E57:P57)", e(ErrorValue::Div0)),
        ("CORREL(E56:P56,E57:P57)", e(ErrorValue::Div0)),
    ]);
}

#[test]
fn rsq_matches_the_microsoft_example_and_correl_pairing() {
    // names fold case like every builtin
    near(eval("rsq(A2:A8,B2:B8)"), 0.05795019157088122, 1e-12, "RSQ");
    check(&[
        ("RSQ({1,2,3},{1,2,3})", n(1.0)),
        ("RSQ({1,2,3},{2,4,6})", n(1.0)),
        ("ROUND(RSQ({1,2,3},{3,6,9}),12)", n(1.0)),
        // one pair, or no spread on one side
        ("RSQ(A2,B2)", e(ErrorValue::Div0)),
        ("RSQ({1,1,1},{1,2,3})", e(ErrorValue::Div0)),
        // ranges of different sizes
        ("RSQ(A2:A8,B2:B7)", e(ErrorValue::NA)),
        ("SLOPE(A2:A8,B2:B7)", e(ErrorValue::NA)),
        ("INTERCEPT(A2:A8,B2:B7)", e(ErrorValue::NA)),
        ("CORREL(A2:A8,B2:B7)", e(ErrorValue::NA)),
    ]);
}
