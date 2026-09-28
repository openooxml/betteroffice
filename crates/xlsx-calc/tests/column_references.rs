use std::cell::Cell as Counter;

use xlsx_calc::{ColumnRange, EvalContext, Expr, evaluate, parse_formula, references};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{
    Cell, CellProvider, CellRef, CellValue, DefinedName, ErrorValue, Sheet, SheetId, Workbook,
};

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

/// One holds only A1:A3; Two holds A1:B10 and, in C, a column of numbers as
/// tall as a row is wide.
fn two_sheets() -> Workbook {
    let mut one = Sheet::new("One");
    for (row, value) in [1.0, 2.0, 4.0].into_iter().enumerate() {
        one.set_cell(CellRef::new(row as u32, 0), number(value));
    }
    let mut two = Sheet::new("Two");
    for row in 0..10 {
        two.set_cell(CellRef::new(row, 0), number(1.0));
        two.set_cell(CellRef::new(row, 1), number(f64::from(row + 1)));
    }
    for row in 0..MAX_COLS {
        two.set_cell(CellRef::new(row, 2), number(1.0));
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(one);
    workbook.sheets.push(two);
    for (name, formula) in [("Keys", "One!$A:$A"), ("Vals", "Two!$B:$B")] {
        workbook.defined_names.push(DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        });
    }
    workbook
}

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

fn plain(workbook: &Workbook, formula: &str) -> CellValue {
    let context = EvalContext::new(workbook, SheetId(0));
    evaluate(&parse_formula(formula).unwrap(), &context)
}

/// the formula evaluated as written and inside `LET`, whose body takes the
/// array evaluator; both must agree.
fn on_one(workbook: &Workbook, formula: &str) -> CellValue {
    let scalar = plain(workbook, formula);
    assert_eq!(
        plain(workbook, &format!("_xlfn.LET(_xlpm.s,0,{formula})")),
        scalar,
        "array evaluation of {formula}"
    );
    scalar
}

fn text_length(value: CellValue) -> usize {
    match value {
        CellValue::Text { value } => value.chars().count(),
        other => panic!("expected text, got {other:?}"),
    }
}

/// a whole-column or whole-row read stops at the used range, but the blanks
/// past it still count wherever a result depends on them.
#[test]
fn blank_counts_include_the_cells_past_the_used_range() {
    let workbook = two_sheets();
    let blanks = f64::from(MAX_ROWS - 3);
    for (formula, expected) in [
        ("COUNTBLANK(A:A)", blanks),
        ("COUNTBLANK(A1:A1048576)", blanks),
        ("COUNTBLANK(1:1)", f64::from(MAX_COLS - 1)),
        ("COUNTIF(A:A,\"\")", blanks),
        ("COUNTIF(A:A,\"<>5\")", f64::from(MAX_ROWS)),
        ("COUNTIF(A:A,\"<>\")", 3.0),
        ("COUNTIFS(A:A,\"\",B:B,\"\")", blanks),
        ("COUNTIFS(A:A,\">1\",B:B,\"\")", 2.0),
    ] {
        assert_eq!(
            on_one(&workbook, formula),
            CellValue::Number { value: expected },
            "{formula}"
        );
    }
    let cols = MAX_COLS as usize;
    for (formula, length) in [
        ("TEXTJOIN(\",\",FALSE,1:1)", cols),
        ("TEXTJOIN(\",\",FALSE,Two!1:1)", cols + 2),
        ("TEXTJOIN(\"\",FALSE,Two!1:1)", 3),
    ] {
        assert_eq!(text_length(plain(&workbook, formula)), length, "{formula}");
    }
    assert_eq!(
        plain(&workbook, "TEXTJOIN(\",\",FALSE,A:A)"),
        CellValue::Error {
            value: ErrorValue::Value
        }
    );
}

/// cutting a reference to the used range leaves its shape alone: an index
/// may still reach past the data, and ranges of different sizes still differ.
#[test]
fn a_cut_reference_keeps_its_shape() {
    let workbook = two_sheets();
    for (formula, expected) in [
        ("HLOOKUP(1,A:A,5,FALSE)", CellValue::Empty),
        ("VLOOKUP(1,1:1,5,FALSE)", CellValue::Empty),
        (
            "COUNTIFS(A:A,\">1\",B1:B3,\"\")",
            CellValue::Error {
                value: ErrorValue::Value,
            },
        ),
        (
            "SUMIFS(B1:B3,A:A,\">1\")",
            CellValue::Error {
                value: ErrorValue::Value,
            },
        ),
        (
            "MMULT(1:1,Two!C1:C16384)",
            CellValue::Error {
                value: ErrorValue::Value,
            },
        ),
    ] {
        assert_eq!(on_one(&workbook, formula), expected, "{formula}");
    }
}

/// whole references on two sheets are cut to one extent, so criteria and
/// the values they pick, or the ranges a product pairs, still line up row for
/// row past the shorter sheet's data.
#[test]
fn whole_references_on_two_sheets_stay_aligned() {
    let workbook = two_sheets();
    for (formula, expected) in [
        ("SUMIF(A:A,\"\",Two!B:B)", 49.0),
        ("SUMIF(A:A,\"\",Two!B1)", 49.0),
        ("SUMIFS(Two!B:B,A:A,\"\")", 49.0),
        ("AVERAGEIF(A:A,\"\",Two!B:B)", 7.0),
        ("AVERAGEIF(A:A,\"\",Two!B1)", 7.0),
        ("MAXIFS(Two!B:B,A:A,\"\")", 10.0),
    ] {
        assert_eq!(
            on_one(&workbook, formula),
            CellValue::Number { value: expected },
            "{formula}"
        );
    }
    for formula in [
        "SUMPRODUCT(A:A,Two!B:B)",
        "SUMPRODUCT(Two!B:B,A:A)",
        "SUMPRODUCT(Keys,Vals)",
        "SUMPRODUCT(Vals,Keys)",
    ] {
        assert_eq!(
            plain(&workbook, formula),
            CellValue::Number { value: 17.0 },
            "{formula}"
        );
    }
    let CellValue::Number { value } = plain(&workbook, "CORREL(A:A,Two!B:B)") else {
        panic!("CORREL pairs the three rows One holds");
    };
    assert!((value - 3.0 / (42.0_f64 / 9.0 * 2.0).sqrt()).abs() < 1e-12);
}

/// references cut to one extent that still spans every column are read as
/// that rectangle, not cut again to each sheet's own width.
#[test]
fn a_shared_cut_is_read_without_cutting_again() {
    let mut one = Sheet::new("One");
    one.set_cell(CellRef::new(0, 0), number(2.0));
    one.set_cell(CellRef::new(1, 0), number(3.0));
    let mut two = Sheet::new("Two");
    two.set_cell(CellRef::new(0, 1), number(7.0));
    two.set_cell(CellRef::new(0, MAX_COLS - 1), number(1.0));
    let mut workbook = Workbook::default();
    workbook.sheets.push(one);
    workbook.sheets.push(two);
    for formula in ["SUMPRODUCT(One!1:2,Two!1:2)", "SUMPRODUCT(Two!1:2,One!1:2)"] {
        assert_eq!(
            plain(&workbook, formula),
            CellValue::Number { value: 0.0 },
            "{formula}"
        );
    }
}

/// inside array evaluation a whole reference reads as its used range, so a
/// column taken from a computed block pairs with another computed from the
/// same reference.
#[test]
fn array_evaluation_reads_whole_references_as_their_used_range() {
    let workbook = two_sheets();
    assert_eq!(
        on_one(&workbook, "SUMPRODUCT(INDEX(A:A+0,0,1),A:A+0)"),
        CellValue::Number { value: 21.0 }
    );
}

/// whole references are cut where each sheet's data ends, not where the
/// furthest sheet the formula names does, so a branch never taken or a range
/// on another sheet costs its own operands nothing.
#[test]
fn a_far_sheet_does_not_stretch_other_references() {
    let mut one = Sheet::new("One");
    for (row, value) in [1.0, 2.0, 4.0].into_iter().enumerate() {
        one.set_cell(CellRef::new(row as u32, 0), number(value));
    }
    one.set_cell(CellRef::new(0, 1), number(10.0));
    let mut two = Sheet::new("Two");
    two.set_cell(CellRef::new(599_999, 0), number(1.0));
    let mut other = Sheet::new("Other");
    other.set_cell(CellRef::new(MAX_ROWS - 1, 0), number(1.0));
    let mut workbook = Workbook::default();
    workbook.sheets.push(one);
    workbook.sheets.push(two);
    workbook.sheets.push(other);
    for (formula, expected) in [
        ("IF(TRUE,SUMPRODUCT(One!A:A,One!A:A),SUM(Two!A:A))", 21.0),
        (
            "_xlfn.LET(_xlpm.x,IF(FALSE,Other!A:A,0),SUM(One!A:A,One!B:B))",
            17.0,
        ),
        ("SUMPRODUCT(One!A:A,One!A:A)+SUM(Two!A:A)", 22.0),
    ] {
        assert_eq!(
            on_one(&workbook, formula),
            CellValue::Number { value: expected },
            "{formula}"
        );
    }
}
