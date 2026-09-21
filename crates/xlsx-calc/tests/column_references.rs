use std::cell::Cell as Counter;

use xlsx_calc::{ColumnRange, EvalContext, Expr, evaluate, parse_formula, references};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellProvider, CellRef, CellValue, ErrorValue, SheetId};

#[test]
fn parses_and_prints_anchored_column_ranges() {
    for (source, expected) in [
        ("S:V", "S:V"),
        ("$S:$V", "$S:$V"),
        ("S:$V", "S:$V"),
        ("$S:V", "$S:V"),
        ("$V:S", "S:$V"),
        ("a : $b", "A:$B"),
        ("A:A", "A:A"),
        ("A:XFD", "A:XFD"),
        ("Sheet1!S:V", "Sheet1!S:V"),
        ("'Other Sheet'!$S:$V", "'Other Sheet'!$S:$V"),
        ("'O''Brien'!S:$V", "'O''Brien'!S:$V"),
    ] {
        let expression = parse_formula(source).unwrap();
        assert!(matches!(expression, Expr::ColumnRange { .. }), "{source}");
        assert_eq!(expression.to_formula(), expected);
        assert_eq!(parse_formula(expected).unwrap(), expression);
    }
    assert!(matches!(parse_formula("S").unwrap(), Expr::Name { .. }));
    assert!(matches!(parse_formula("V1").unwrap(), Expr::Ref { .. }));
    assert!(matches!(parse_formula("1E3").unwrap(), Expr::Number(_)));
    assert!(matches!(
        parse_formula("LOG10(A1)").unwrap(),
        Expr::FuncCall { .. }
    ));
}

#[test]
fn rejects_invalid_column_ranges() {
    for source in [
        "A:", ":B", "A:B1", "A1:B", "A:1", "$:B", "A$:B", "XFE:XFF", "A:B:C",
    ] {
        assert!(parse_formula(source).is_err(), "{source}");
    }
}

#[test]
fn full_height_dependencies_do_not_allocate_cells() {
    let expression = parse_formula("VLOOKUP(A1,Data!$S:$V,4,FALSE)").unwrap();
    let dependencies = references(&expression);
    assert_eq!(dependencies.len(), 2);
    assert_eq!(dependencies[1].0.as_deref(), Some("Data"));
    assert_eq!(dependencies[1].1.to_a1(), "S1:V1048576");
    let range = ColumnRange::parse_a1("A:XFD").unwrap().cell_range();
    assert!(range.contains(CellRef::new(MAX_ROWS - 1, MAX_COLS - 1)));
}

struct CountedData {
    visits: Counter<usize>,
}

impl CellProvider for CountedData {
    fn value(&self, _sheet: SheetId, at: CellRef) -> CellValue {
        self.visits.set(self.visits.get() + 1);
        match (at.row, at.col) {
            (1, 18) => CellValue::Number { value: 2.0 },
            (1, 21) => CellValue::Text {
                value: "two".into(),
            },
            (2, 18) => CellValue::Number { value: 3.0 },
            (2, 21) => CellValue::Text {
                value: "three".into(),
            },
            (row, 21) if row == MAX_ROWS - 1 => CellValue::Number { value: 42.0 },
            _ => CellValue::Empty,
        }
    }

    fn formula(&self, _sheet: SheetId, _at: CellRef) -> Option<&str> {
        None
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        name.eq_ignore_ascii_case("Data").then_some(SheetId(0))
    }

    fn used_rows(&self, _sheet: SheetId) -> u32 {
        MAX_ROWS
    }

    fn used_cols(&self, _sheet: SheetId) -> u32 {
        MAX_COLS
    }
}

#[test]
fn an_approximate_lookup_reads_past_an_out_of_order_header() {
    struct Listed;
    impl CellProvider for Listed {
        fn value(&self, _sheet: SheetId, at: CellRef) -> CellValue {
            match (at.col, at.row) {
                (0, 0) => CellValue::Text { value: "ID".into() },
                (0, row) if row <= 4 => CellValue::Number { value: row as f64 },
                (1, row) if (1..=4).contains(&row) => CellValue::Number {
                    value: row as f64 * 10.0,
                },
                _ => CellValue::Empty,
            }
        }

        fn formula(&self, _sheet: SheetId, _at: CellRef) -> Option<&str> {
            None
        }

        fn sheet_id(&self, _name: &str) -> Option<SheetId> {
            None
        }

        fn used_rows(&self, _sheet: SheetId) -> u32 {
            5
        }
    }
    let context = EvalContext::new(&Listed, SheetId(0));
    for (formula, expected) in [
        ("VLOOKUP(1,A1:B5,2)", 10.0),
        ("VLOOKUP(3,A1:B5,2)", 30.0),
        ("VLOOKUP(3.5,A1:B5,2)", 30.0),
        ("VLOOKUP(9,A:B,2)", 40.0),
    ] {
        assert_eq!(
            evaluate(&parse_formula(formula).unwrap(), &context),
            CellValue::Number { value: expected },
            "{formula}"
        );
    }
}

#[test]
fn criteria_over_a_whole_column_stop_at_the_used_range() {
    struct Sparse;
    impl CellProvider for Sparse {
        fn value(&self, _sheet: SheetId, at: CellRef) -> CellValue {
            match (at.col, at.row) {
                (0, 0) => CellValue::Number { value: 1.0 },
                (0, 1) => CellValue::Number { value: 3.0 },
                (0, 2) => CellValue::Number { value: 5.0 },
                _ => CellValue::Empty,
            }
        }

        fn formula(&self, _sheet: SheetId, _at: CellRef) -> Option<&str> {
            None
        }

        fn sheet_id(&self, _name: &str) -> Option<SheetId> {
            None
        }

        fn used_rows(&self, _sheet: SheetId) -> u32 {
            3
        }
    }
    let context = EvalContext::new(&Sparse, SheetId(0));
    for (formula, expected) in [
        ("COUNTIF(A:A,\">2\")", 2.0),
        ("SUMIF(A:A,\">2\")", 8.0),
        ("COUNTIFS(A:A,\">2\",B:B,\"\")", 2.0),
        ("COUNT(A:A)", 3.0),
        ("SUM(A:A)", 9.0),
    ] {
        assert_eq!(
            evaluate(&parse_formula(formula).unwrap(), &context),
            CellValue::Number { value: expected },
            "{formula}"
        );
    }
}

#[test]
fn lookups_and_metadata_do_not_materialize_entire_columns() {
    for (formula, expected, visits) in [
        (
            "VLOOKUP(2,S:XFD,4,FALSE)",
            CellValue::Text {
                value: "two".into(),
            },
            3,
        ),
        (
            "VLOOKUP(2.5,S2:V3,4,TRUE)",
            CellValue::Text {
                value: "two".into(),
            },
            3,
        ),
        (
            "ROWS(S:V)",
            CellValue::Number {
                value: MAX_ROWS as f64,
            },
            0,
        ),
        (
            "COLUMNS(A:XFD)",
            CellValue::Number {
                value: MAX_COLS as f64,
            },
            0,
        ),
        ("INDEX(S:V,1048576,4)", CellValue::Number { value: 42.0 }, 1),
        ("VLOOKUP(0,S:V,4,FALSE)", CellValue::Empty, 2),
        (
            "VLOOKUP(42,V:X,1,FALSE)",
            CellValue::Number { value: 42.0 },
            MAX_ROWS as usize + 1,
        ),
        (
            "SUM(A:XFD)",
            CellValue::Error {
                value: ErrorValue::Num,
            },
            0,
        ),
        (
            "XLOOKUP(2,Data!S:S,Data!V:V)",
            CellValue::Text {
                value: "two".into(),
            },
            3,
        ),
    ] {
        let data = CountedData {
            visits: Counter::new(0),
        };
        let context = EvalContext::new(&data, SheetId(0));
        assert_eq!(
            evaluate(&parse_formula(formula).unwrap(), &context),
            expected,
            "{formula}"
        );
        assert_eq!(data.visits.get(), visits, "{formula}");
    }
}
