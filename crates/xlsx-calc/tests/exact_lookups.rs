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

fn b(value: bool) -> CellValue {
    CellValue::Bool { value }
}

fn na() -> CellValue {
    CellValue::Error {
        value: ErrorValue::NA,
    }
}

fn set(sheet: &mut Sheet, at: &str, value: CellValue) {
    sheet.set_cell(
        CellRef::parse_a1(at).unwrap(),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

/// the model sheet of a cost workbook: A2 is blank above the row labels and
/// B2:D2 name the scenarios, so a scenario key of 0 has no column to land on.
/// F1:F5 hold a blank, a real "" text, 0, FALSE and "x" against G 101..105;
/// H1:H2 hold a blank and "a" against I 201..202.
fn workbook() -> Workbook {
    let mut sheet = Sheet::new("Model");
    for (at, value) in [
        ("B2", t("Low")),
        ("C2", t("Central")),
        ("D2", t("High")),
        ("A3", t("Used Evidence pages")),
        ("B3", n(1.0)),
        ("C3", n(2.0)),
        ("D3", n(3.0)),
        ("A4", t("Mins per used page")),
        ("B4", n(0.25)),
        ("C4", n(1.0)),
        ("D4", n(1.5)),
        ("B62", n(0.0)),
        ("F2", t("")),
        ("F3", n(0.0)),
        ("F4", b(false)),
        ("F5", t("x")),
        ("H2", t("a")),
        ("I1", n(201.0)),
        ("I2", n(202.0)),
    ] {
        set(&mut sheet, at, value);
    }
    for row in 1..=5 {
        set(&mut sheet, &format!("G{row}"), n(100.0 + row as f64));
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    workbook
}

fn check(cases: &[(&str, CellValue)]) {
    let workbook = workbook();
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
fn match_zero_skips_a_leading_blank_inside_offset() {
    check(&[
        (
            "OFFSET($A$2,MATCH(\"Used Evidence pages\",$A$2:$A$12,0)-1,MATCH($B$62,$A$2:$D$2,0)-1)",
            na(),
        ),
        ("MATCH($B$62,$A$2:$D$2,0)", na()),
        (
            "OFFSET($A$2,MATCH(\"Used Evidence pages\",$A$2:$A$12,0)-1,MATCH(\"Central\",$A$2:$D$2,0)-1)",
            n(2.0),
        ),
    ]);
}

#[test]
fn exact_match_never_lands_on_a_blank_cell() {
    check(&[
        ("MATCH(0,H1:H2,0)", na()),
        ("MATCH(\"\",H1:H2,0)", na()),
        ("MATCH(FALSE,H1:H2,0)", na()),
        ("MATCH(\"a\",H1:H2,0)", n(2.0)),
        ("MATCH(0,F1:F5,0)", n(3.0)),
        ("MATCH(\"\",F1:F5,0)", n(2.0)),
        ("MATCH(FALSE,F1:F5,0)", n(4.0)),
    ]);
}

#[test]
fn exact_vlookup_and_hlookup_skip_blank_keys() {
    check(&[
        ("VLOOKUP(0,H1:I2,2,FALSE)", na()),
        ("VLOOKUP(\"\",H1:I2,2,FALSE)", na()),
        ("VLOOKUP(0,F1:G5,2,FALSE)", n(103.0)),
        ("VLOOKUP(\"\",F1:G5,2,0)", n(102.0)),
        ("VLOOKUP(FALSE,F1:G5,2,0)", n(104.0)),
        ("HLOOKUP(0,A2:D3,2,FALSE)", na()),
        ("HLOOKUP(\"High\",A2:D3,2,FALSE)", n(3.0)),
    ]);
}

#[test]
fn xmatch_and_xlookup_exact_skip_blank_cells() {
    check(&[
        ("XMATCH(0,H1:H2)", na()),
        ("XMATCH(\"\",H1:H2)", na()),
        ("XMATCH(0,H1:H2,0,-1)", na()),
        ("XMATCH(\"\",F1:F5)", n(2.0)),
        ("XMATCH(0,F1:F5,0,-1)", n(3.0)),
        ("XMATCH(0,$A$2:$D$2)", na()),
        ("XLOOKUP(0,H1:H2,I1:I2)", na()),
        ("XLOOKUP(FALSE,H1:H2,I1:I2,\"none\")", t("none")),
        ("XLOOKUP(\"\",F1:F5,G1:G5)", n(102.0)),
        ("XLOOKUP(FALSE,F1:F5,G1:G5)", n(104.0)),
    ]);
}

#[test]
fn array_keys_skip_blank_cells_too() {
    let workbook = workbook();
    for (formula, expected) in [
        ("COUNT(MATCH({0;\"a\"},H1:H2,0))", n(1.0)),
        ("COUNT(XMATCH({0;\"a\"},H1:H2))", n(1.0)),
        (
            "SUM(IFERROR(XLOOKUP({0;\"\";\"a\"},H1:H2,I1:I2),0))",
            n(202.0),
        ),
        ("SUM(MATCH({0;\"\"},F1:F5,0))", n(5.0)),
    ] {
        let expression = parse_formula(formula).unwrap();
        assert_eq!(
            evaluate_array(&expression, &EvalContext::new(&workbook, SheetId(0))).into_scalar(),
            expected,
            "array: {formula}"
        );
    }
}
