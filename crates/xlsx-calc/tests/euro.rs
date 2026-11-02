//! EUROCONVERT, checked against Microsoft's documented examples and 30-digit
//! mpmath oracles on the fixed conversion rates.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

/// the value `formula` takes in F1 of a sheet holding `inputs`.
fn eval_in(inputs: &[(&str, CellValue)], formula: &str) -> CellValue {
    let mut sheet = Sheet::new("Sheet1");
    for (address, value) in inputs {
        sheet.set_cell(
            CellRef::parse_a1(address).unwrap(),
            Cell {
                value: value.clone(),
                ..Cell::default()
            },
        );
    }
    let at = CellRef::parse_a1("F1").unwrap();
    sheet.set_cell(
        at,
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    rebuild_and_recalc_all(&mut workbook, None);
    workbook.value(SheetId(0), at)
}

fn eval(formula: &str) -> CellValue {
    eval_in(&[], formula)
}

fn close(formula: &str, want: f64) {
    match eval(formula) {
        CellValue::Number { value } => assert!(
            (value - want).abs() <= 1e-12 * want.abs().max(1.0),
            "{formula}: {value} vs {want}"
        ),
        other => panic!("{formula}: {other:?}"),
    }
}

#[test]
fn euroconvert_microsoft_examples() {
    let table = [
        ("A2", number(1.2)),
        ("B2", text("DEM")),
        ("C2", text("EUR")),
        ("A3", number(1.0)),
        ("B3", text("FRF")),
        ("C3", text("EUR")),
        ("A5", number(1.0)),
        ("B5", text("FRF")),
        ("C5", text("DEM")),
    ];
    assert_eq!(eval_in(&table, "EUROCONVERT(A2,B2,C2)"), number(0.61));
    assert_eq!(
        eval_in(&table, "EUROCONVERT(A3,B3,C3,TRUE,3)"),
        number(0.152)
    );
    assert_eq!(
        eval_in(&table, "EUROCONVERT(A3,B3,C3,FALSE,3)"),
        number(0.15)
    );
    match eval_in(&table, "EUROCONVERT(A5,B5,C5,TRUE,3)") {
        CellValue::Number { value } => assert!((value - 0.29728616).abs() < 1e-15),
        other => panic!("{other:?}"),
    }
}

/// mpmath 30 digits on the fixed rates.
#[test]
fn euroconvert_matches_mpmath_oracles() {
    close(
        "EUROCONVERT(100,\"DEM\",\"FRF\",TRUE)",
        335.385_488_513_827_9,
    );
    close(
        "EUROCONVERT(1000,\"ITL\",\"EUR\",TRUE)",
        0.516_456_899_089_486_5,
    );
    close("EUROCONVERT(1000,\"FRF\",\"DEM\",TRUE,3)", 298.164_327_67);
    close("EUROCONVERT(1000,\"FRF\",\"DEM\",FALSE,3)", 298.16);
    close("EUROCONVERT(250,\"EUR\",\"GRD\")", 85188.0);
    close("EUROCONVERT(50,\"IEP\",\"PTE\")", 12728.0);
    close("EUROCONVERT(-1234.5,\"ATS\",\"EUR\")", -89.71);
}

#[test]
fn euroconvert_same_currency_returns_the_amount_unrounded() {
    close("EUROCONVERT(1.23456,\"DEM\",\"dem\")", 1.23456);
    close("EUROCONVERT(1000,\"BEF\",\"LUF\")", 1000.0);
    close(
        "EUROCONVERT(12.3456,\"eur\",\"Frf\",TRUE)",
        12.3456 * 6.55957,
    );
}

#[test]
fn euroconvert_invalid_arguments_are_value_errors() {
    let value = CellValue::Error {
        value: ErrorValue::Value,
    };
    for formula in [
        "EUROCONVERT(1,\"USD\",\"EUR\")",
        "EUROCONVERT(1,\"DEM\",\"XEU\")",
        "EUROCONVERT(1,\"DEM\",\"EUR\",TRUE,2)",
        "EUROCONVERT(1,\"DEM\",\"EUR\",TRUE,2.9)",
        "EUROCONVERT(\"abc\",\"DEM\",\"EUR\")",
        "EUROCONVERT(1,\"DEM\",\"EUR\",\"maybe\")",
        "EUROCONVERT(1,\"DEM\")",
        "EUROCONVERT(1,\"DEM\",\"EUR\",TRUE,3,1)",
    ] {
        assert_eq!(eval(formula), value, "{formula}");
    }
    assert_eq!(
        eval("EUROCONVERT(1/0,\"DEM\",\"EUR\")"),
        CellValue::Error {
            value: ErrorValue::Div0
        }
    );
}

/// amounts too large to carry the rounding's decimals are already whole at
/// that precision, so rounding them keeps the finite result.
#[test]
fn euroconvert_rounding_of_huge_amounts_stays_finite() {
    close("EUROCONVERT(1E307,\"DEM\",\"EUR\",TRUE,3)", 1e307 / 1.95583);
    close(
        "EUROCONVERT(1E307,\"DEM\",\"EUR\",FALSE,3)",
        1e307 / 1.95583,
    );
    close("EUROCONVERT(1E307,\"DEM\",\"EUR\")", 1e307 / 1.95583);
    close("EUROCONVERT(1E307,\"EUR\",\"DEM\")", 1e307 * 1.95583);
    close(
        "EUROCONVERT(-1E307,\"FRF\",\"DEM\",FALSE,15)",
        -1e307 / 6.55957 * 1.95583,
    );
    assert_eq!(
        eval("EUROCONVERT(1E308,\"EUR\",\"DEM\")"),
        CellValue::Error {
            value: ErrorValue::Num
        }
    );
}
