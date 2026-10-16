use xlsx_calc::{EvalContext, evaluate, evaluate_array, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn t(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn na() -> CellValue {
    CellValue::Error {
        value: ErrorValue::NA,
    }
}

fn grid(keys: &[CellValue]) -> Workbook {
    let mut sheet = Sheet::new("Version");
    for (i, value) in keys.iter().enumerate() {
        for at in [CellRef::new(i as u32, 0), CellRef::new(0, i as u32 + 2)] {
            sheet.set_cell(
                at,
                Cell {
                    value: value.clone(),
                    ..Cell::default()
                },
            );
        }
        sheet.set_cell(
            CellRef::new(i as u32, 1),
            Cell {
                value: n(i as f64 + 101.0),
                ..Cell::default()
            },
        );
    }
    for at in [CellRef::new(41, 0), CellRef::new(0, 43)] {
        sheet.set_cell(
            at,
            Cell {
                style: Some(0),
                ..Cell::default()
            },
        );
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    workbook
}

fn check(keys: &[CellValue], cases: &[(&str, CellValue)]) {
    let workbook = grid(keys);
    for (formula, expected) in cases {
        let expression = parse_formula(formula).unwrap();
        assert_eq!(
            evaluate(&expression, &EvalContext::new(&workbook, SheetId(0))),
            *expected,
            "scalar: {formula}"
        );
        assert_eq!(
            evaluate_array(&expression, &EvalContext::new(&workbook, SheetId(0))).into_scalar(),
            *expected,
            "array: {formula}"
        );
    }
}

#[test]
fn vlookup_numeric_keys_before_styled_blanks() {
    check(
        &[n(1.0), n(2.0), n(3.0), n(5.0), n(8.7)],
        &[("VLOOKUP(99999,Version!A:A,1)", n(8.7))],
    );
}

#[test]
fn match_ascending_keys_before_styled_blanks() {
    check(&[n(1.0), n(2.0), n(8.7)], &[("MATCH(99999,A:A,1)", n(3.0))]);
}

#[test]
fn approximate_keys_with_interleaved_blanks_keep_original_positions() {
    check(
        &[CellValue::Empty, n(-3.0), CellValue::Empty, n(2.0), n(8.7)],
        &[
            ("VLOOKUP(3,A1:B5,2)", n(104.0)),
            ("HLOOKUP(3,C1:G1,1)", n(2.0)),
            ("MATCH(-1,A1:A5,1)", n(2.0)),
            ("MATCH(3,A1:A5,1)", n(4.0)),
            ("LOOKUP(3,A1:A5,B1:B5)", n(104.0)),
            ("MATCH(3,VSTACK(A1:A5),1)", n(4.0)),
        ],
    );
}

#[test]
fn approximate_text_keys_before_blanks() {
    check(
        &[t("apple"), t("pear")],
        &[
            ("VLOOKUP(\"zebra\",A:A,1)", t("pear")),
            ("HLOOKUP(\"zebra\",C1:AR1,1)", t("pear")),
            ("MATCH(\"zebra\",A:A,1)", n(2.0)),
            ("LOOKUP(\"zebra\",A1:A42)", t("pear")),
        ],
    );
}

#[test]
fn approximate_value_below_first_key_is_na() {
    check(
        &[CellValue::Empty, n(2.0), n(8.7)],
        &[
            ("VLOOKUP(1,A:A,1)", na()),
            ("HLOOKUP(1,C1:AR1,1)", na()),
            ("MATCH(1,A:A,1)", na()),
            ("LOOKUP(1,A1:A42)", na()),
        ],
    );
}

#[test]
fn match_descending_negative_keys_before_blanks() {
    check(
        &[n(-1.0), CellValue::Empty, n(-3.0), n(-8.7)],
        &[
            ("MATCH(-99999,A:A,-1)", n(4.0)),
            ("MATCH(-2,A:A,-1)", n(1.0)),
            ("MATCH(0,A:A,-1)", na()),
            ("MATCH(-99999,VSTACK(A1:A42),-1)", n(4.0)),
        ],
    );
}

#[test]
fn approximate_all_blank_column_is_na() {
    check(
        &[],
        &[
            ("VLOOKUP(99999,A:A,1)", na()),
            ("HLOOKUP(99999,C1:AR1,1)", na()),
            ("MATCH(99999,A:A,1)", na()),
            ("MATCH(-99999,A:A,-1)", na()),
            ("LOOKUP(99999,A1:A42)", na()),
        ],
    );
}

#[test]
fn approximate_lookups_skip_other_key_types() {
    check(
        &[
            t("header"),
            n(2.0),
            t("text"),
            n(8.7),
            CellValue::Bool { value: true },
        ],
        &[
            ("VLOOKUP(3,A:A,1)", n(2.0)),
            ("MATCH(99999,A:A,1)", n(4.0)),
            ("MATCH(3,VSTACK(A1:A5),1)", n(2.0)),
            ("LOOKUP(99999,A1:A42)", n(8.7)),
            ("MATCH(\"zebra\",A:A,1)", n(3.0)),
            ("MATCH(TRUE(),A:A,1)", n(5.0)),
        ],
    );
    check(
        &[t("header"), n(8.7), t("text"), n(2.0)],
        &[
            ("MATCH(3,A:A,-1)", n(2.0)),
            ("MATCH(3,VSTACK(A1:A5),-1)", n(2.0)),
        ],
    );
}

#[test]
fn binary_xmatch_skips_blanks_and_other_types() {
    check(
        &[t("header"), n(-3.0), CellValue::Empty, n(2.0), t("text")],
        &[
            ("XMATCH(-1,A1:A42,-1,2)", n(2.0)),
            ("XMATCH(0,A1:A42,-1,2)", n(2.0)),
        ],
    );
    check(
        &[t("header"), n(2.0), CellValue::Empty, n(-3.0), t("text")],
        &[("XMATCH(-1,A1:A42,1,-2)", n(2.0))],
    );
}

#[test]
fn binary_xlookup_skips_blanks_and_other_types() {
    let workbook = grid(&[n(-3.0), CellValue::Empty, n(2.0), t("text")]);
    for (formula, expected) in [
        ("XLOOKUP(-1,A1:A42,B1:B42,,-1,2)", n(101.0)),
        ("XLOOKUP(0,A1:A42,B1:B42,,-1,2)", n(101.0)),
        ("XLOOKUP(3,A1:A42,B1:B42,,1,2)", na()),
    ] {
        assert_eq!(
            evaluate_array(
                &parse_formula(formula).unwrap(),
                &EvalContext::new(&workbook, SheetId(0))
            )
            .into_scalar(),
            expected,
            "{formula}"
        );
    }
    let workbook = grid(&[n(2.0), CellValue::Empty, n(-3.0), t("text")]);
    assert_eq!(
        evaluate_array(
            &parse_formula("XLOOKUP(-1,A1:A42,B1:B42,,1,-2)").unwrap(),
            &EvalContext::new(&workbook, SheetId(0))
        )
        .into_scalar(),
        n(101.0)
    );
}

#[test]
fn exact_lookup_blank_coercion_is_unchanged() {
    check(
        &[CellValue::Empty, n(2.0)],
        &[
            ("VLOOKUP(0,A1:B2,2,FALSE)", n(101.0)),
            ("HLOOKUP(0,C1:D1,1,FALSE)", CellValue::Empty),
            ("MATCH(0,A1:A2,0)", n(1.0)),
            ("XMATCH(0,A1:A2,0,2)", n(1.0)),
            ("XLOOKUP(0,A1:A2,B1:B2)", n(101.0)),
        ],
    );
}
