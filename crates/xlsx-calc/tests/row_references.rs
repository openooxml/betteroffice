//! `2:7`, a reference to every column of a span of rows.

use xlsx_calc::{EvalContext, Expr, RowRange, Value, evaluate_array, parse_formula, references};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{Cell, CellRef, CellValue, Sheet, SheetId, Workbook};

#[test]
fn parses_and_prints_anchored_row_ranges() {
    for (source, expected) in [
        ("2:7", "2:7"),
        ("$2:$7", "$2:$7"),
        ("2:$7", "2:$7"),
        ("$2:7", "$2:7"),
        ("$7:2", "2:$7"),
        ("1 : $3", "1:$3"),
        ("1:1", "1:1"),
        ("1:1048576", "1:1048576"),
        ("Sheet1!2:7", "Sheet1!2:7"),
        ("'Other Sheet'!$2:$7", "'Other Sheet'!$2:$7"),
    ] {
        let expression = parse_formula(source).unwrap();
        assert!(matches!(expression, Expr::RowRange { .. }), "{source}");
        assert_eq!(expression.to_formula(), expected);
        assert_eq!(parse_formula(expected).unwrap(), expression);
    }
    assert!(matches!(parse_formula("1E3").unwrap(), Expr::Number(_)));
    assert!(matches!(parse_formula("12").unwrap(), Expr::Number(_)));
    assert!(matches!(
        parse_formula("A:B").unwrap(),
        Expr::ColumnRange { .. }
    ));
}

#[test]
fn rejects_invalid_row_ranges() {
    for source in [
        "1:",
        ":2",
        "1:B2",
        "A1:2",
        "1:A",
        "$:2",
        "1048577:1048577",
        "1:2:3",
    ] {
        assert!(parse_formula(source).is_err(), "{source}");
    }
}

#[test]
fn a_row_range_depends_on_the_whole_row() {
    let expression = parse_formula("SUM(Data!$3:$4)").unwrap();
    let dependencies = references(&expression);
    assert_eq!(dependencies[0].0.as_deref(), Some("Data"));
    assert_eq!(dependencies[0].1.to_a1(), "A3:XFD4");
    let range = RowRange::parse_a1("1:1048576").unwrap().cell_range();
    assert!(range.contains(CellRef::new(MAX_ROWS - 1, MAX_COLS - 1)));
}

fn fixture() -> Workbook {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    for (index, value) in [5.0, 6.0, 7.0].iter().enumerate() {
        sheet.set_cell(
            CellRef::new(1, index as u32),
            Cell {
                value: CellValue::Number { value: *value },
                ..Cell::default()
            },
        );
    }
    workbook.sheets.push(sheet);
    workbook
}

/// a whole-row block costs the columns the sheet uses, not all 16384.
#[test]
fn a_row_range_materializes_only_the_used_columns() {
    let workbook = fixture();
    let context = EvalContext::new(&workbook, SheetId(0));
    let Value::Array(array) = evaluate_array(&parse_formula("2:2").unwrap(), &context) else {
        panic!("expected a block");
    };
    assert_eq!((array.rows(), array.cols()), (1, 3));
    assert_eq!(array.at(0, 2), CellValue::Number { value: 7.0 });
}

/// `ROW($1:$99)` is the classic CSE idiom for "the numbers 1 to 99".
#[test]
fn row_over_a_row_range_counts_down_the_rows() {
    let workbook = fixture();
    let context = EvalContext::new(&workbook, SheetId(0));
    let Value::Array(array) = evaluate_array(&parse_formula("ROW($1:$4)").unwrap(), &context)
    else {
        panic!("expected a block");
    };
    assert_eq!((array.rows(), array.cols()), (4, 1));
    assert_eq!(array.at(3, 0), CellValue::Number { value: 4.0 });
}

#[test]
fn a_scalar_function_lifts_over_a_row_range_index() {
    let workbook = fixture();
    let context = EvalContext::new(&workbook, SheetId(0));
    let Value::Array(array) = evaluate_array(
        &parse_formula(r#"MID("abc",ROW(1:3),1)"#).unwrap(),
        &context,
    ) else {
        panic!("expected a block");
    };
    assert_eq!(
        (0..3).map(|row| array.at(row, 0)).collect::<Vec<_>>(),
        vec![
            CellValue::Text { value: "a".into() },
            CellValue::Text { value: "b".into() },
            CellValue::Text { value: "c".into() },
        ]
    );
}
