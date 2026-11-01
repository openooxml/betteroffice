//! the cash-flow functions and CELL inside array formulas: scalar arguments
//! lift element by element and broadcast, range arguments stay whole.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRange, CellRef, CellValue, ErrorValue, Sheet, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

/// the example cash flows and dates in A1:B5, and B7 holding 1.
fn inputs() -> Vec<(&'static str, f64)> {
    vec![
        ("A1", -10_000.0),
        ("A2", 2_750.0),
        ("A3", 4_250.0),
        ("A4", 3_250.0),
        ("A5", 2_750.0),
        ("B1", 39_448.0),
        ("B2", 39_508.0),
        ("B3", 39_751.0),
        ("B4", 39_859.0),
        ("B5", 39_904.0),
        ("C1", 100.0),
        ("C2", 200.0),
        ("C3", 300.0),
    ]
}

fn sheet_with_inputs() -> Sheet {
    let mut sheet = Sheet::new("Sheet1");
    for (address, value) in inputs() {
        sheet.set_cell(
            CellRef::parse_a1(address).unwrap(),
            Cell {
                value: number(value),
                ..Cell::default()
            },
        );
    }
    sheet
}

/// the cells an array formula over `reference` fills, row-major.
fn array_formula(formula: &str, reference: &str) -> Vec<CellValue> {
    let range = CellRange::parse_a1(reference).unwrap();
    let mut sheet = sheet_with_inputs();
    sheet.set_cell(
        range.start,
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    sheet.set_array_formula(range.start, range);
    let mut book = Workbook {
        sheets: vec![sheet],
        ..Workbook::default()
    };
    rebuild_and_recalc_all(&mut book, None);
    let mut out = Vec::new();
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            out.push(
                book.sheets[0]
                    .cell(CellRef::new(row, col))
                    .map(|cell| cell.value.clone())
                    .unwrap_or(CellValue::Empty),
            );
        }
    }
    out
}

/// a plain formula's value, for comparing one element against.
fn scalar(formula: &str) -> CellValue {
    array_formula(formula, "Z1")[0].clone()
}

#[test]
fn scalar_functions_lift_over_array_literals() {
    assert_eq!(
        array_formula("SLN({100,200},0,10)", "Z1:AA1"),
        vec![number(10.0), number(20.0)]
    );
    assert_eq!(
        array_formula("SYD(30000,7500,10,{1;10})", "Z1:Z2"),
        vec![
            scalar("SYD(30000,7500,10,1)"),
            scalar("SYD(30000,7500,10,10)")
        ]
    );
    assert_eq!(
        array_formula("RATE({48,60},-200,8000)", "Z1:AA1"),
        vec![scalar("RATE(48,-200,8000)"), scalar("RATE(60,-200,8000)")]
    );
    assert_eq!(
        array_formula("DOLLARDE({1.02,1.1},{16,32})", "Z1:AA1"),
        vec![scalar("DOLLARDE(1.02,16)"), scalar("DOLLARDE(1.1,32)")]
    );
}

#[test]
fn scalar_functions_lift_over_ranges() {
    assert_eq!(
        array_formula("SLN(C1:C3,0,10)", "Z1:Z3"),
        vec![number(10.0), number(20.0), number(30.0)]
    );
    assert_eq!(
        array_formula("IPMT(0.1,{1,2,3},3,C1:C3)", "Z1:AB3"),
        (1..=3)
            .flat_map(|pv| {
                (1..=3).map(move |per| scalar(&format!("IPMT(0.1,{per},3,{})", pv * 100)))
            })
            .collect::<Vec<_>>()
    );
}

#[test]
fn a_scalar_broadcasts_against_an_array() {
    assert_eq!(
        array_formula("SLN(1000,0,{5;10})", "Z1:Z2"),
        vec![number(200.0), number(100.0)]
    );
    assert_eq!(
        array_formula("EFFECT(0.12,{1,4,12})", "Z1:AB1"),
        vec![
            scalar("EFFECT(0.12,1)"),
            scalar("EFFECT(0.12,4)"),
            scalar("EFFECT(0.12,12)"),
        ]
    );
}

/// shorter arrays pad with `#N/A`, as excel pads every lifted function.
#[test]
fn mismatched_shapes_pad_with_na() {
    assert_eq!(
        array_formula("SLN({100,200,300},0,{10,20})", "Z1:AB1"),
        vec![number(10.0), number(10.0), error(ErrorValue::NA)]
    );
}

#[test]
fn errors_stay_with_their_element() {
    assert_eq!(
        array_formula("SLN({100,200},0,{10,0})", "Z1:AA1"),
        vec![number(10.0), error(ErrorValue::Div0)]
    );
    assert_eq!(
        array_formula("DDB(2400,300,10,{1,11})", "Z1:AA1"),
        vec![scalar("DDB(2400,300,10,1)"), error(ErrorValue::Num)]
    );
}

/// cash flows, dates and schedules are read whole; only the scalar
/// arguments around them lift.
#[test]
fn range_arguments_stay_whole() {
    assert_eq!(
        array_formula("XNPV({0.09,0.1},A1:A5,B1:B5)", "Z1:AA1"),
        vec![
            scalar("XNPV(0.09,A1:A5,B1:B5)"),
            scalar("XNPV(0.1,A1:A5,B1:B5)")
        ]
    );
    assert_eq!(
        array_formula("XIRR(A1:A5,B1:B5,{0.1,0.3})", "Z1:AA1"),
        vec![
            scalar("XIRR(A1:A5,B1:B5,0.1)"),
            scalar("XIRR(A1:A5,B1:B5,0.3)")
        ]
    );
    assert_eq!(
        array_formula("IRR(A1:A5,{0.1;0.2})", "Z1:Z2"),
        vec![scalar("IRR(A1:A5,0.1)"), scalar("IRR(A1:A5,0.2)")]
    );
    assert_eq!(
        array_formula("MIRR(A1:A5,{0.1,0.2},0.12)", "Z1:AA1"),
        vec![
            scalar("MIRR(A1:A5,0.1,0.12)"),
            scalar("MIRR(A1:A5,0.2,0.12)")
        ]
    );
    assert_eq!(
        array_formula("FVSCHEDULE({1,2},{0.09,0.11,0.1})", "Z1:AA1"),
        vec![
            scalar("FVSCHEDULE(1,{0.09,0.11,0.1})"),
            scalar("FVSCHEDULE(2,{0.09,0.11,0.1})")
        ]
    );
}

#[test]
fn cell_lifts_its_info_type_and_keeps_the_reference() {
    assert_eq!(
        array_formula("CELL({\"row\",\"col\"},C3)", "Z1:AA1"),
        vec![number(3.0), number(3.0)]
    );
    assert_eq!(
        array_formula("CELL({\"contents\";\"address\"},C2)", "Z1:Z2"),
        vec![
            number(200.0),
            CellValue::Text {
                value: "$C$2".into()
            }
        ]
    );
}
