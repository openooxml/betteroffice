use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn t(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn set(sheet: &mut Sheet, address: &str, value: CellValue) {
    sheet.set_cell(
        CellRef::parse_a1(address).unwrap(),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn workbook() -> Workbook {
    let mut wb = Workbook::default();
    let mut calc = Sheet::new("Calculation_HIDE");
    set(&mut calc, "A4", t("FIRE1202 2019"));
    set(&mut calc, "A3", t("England"));
    set(&mut calc, "A8", n(7.0));
    set(&mut calc, "B2", n(2.0));
    set(&mut calc, "B5", n(11.0));
    set(&mut calc, "E3", n(13.0));
    set(&mut calc, "C6", n(17.0));
    set(&mut calc, "S8", t("'Main data'!R12C3"));
    let mut data = Sheet::new("FIRE1202 2019");
    set(&mut data, "A1", n(7.0));
    set(&mut data, "B1", n(23.0));
    let mut england = Sheet::new("(England)");
    set(&mut england, "D6", n(29.0));
    let mut main = Sheet::new("Main data");
    set(&mut main, "C12", n(31.0));
    let mut apostrophe = Sheet::new("O'Brien!");
    set(&mut apostrophe, "A1", n(37.0));
    wb.sheets = vec![calc, data, england, main, apostrophe];
    wb
}

fn check(cases: &[(&str, CellValue)]) {
    for (formula, expected) in cases {
        let mut wb = workbook();
        let at = CellRef::parse_a1("C4").unwrap();
        wb.sheets[0].set_cell(
            at,
            Cell {
                formula: Some((*formula).into()),
                ..Cell::default()
            },
        );
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(wb.value(SheetId(0), at), *expected, "{formula}");
    }
}

#[test]
fn lowercase_a1_and_case_insensitive_sheets_recalculate() {
    check(&[
        ("e3", n(13.0)),
        ("SUM(b5:c6)", n(28.0)),
        ("INDIRECT(\"e3\")", n(13.0)),
        ("'fire1202 2019'!b1", n(23.0)),
        ("SUM('fire1202 2019'!a1:b1)", n(30.0)),
        ("'Fire1202 2019'!$b$1", n(23.0)),
        ("INDIRECT(\"'\"&$A$4&\"'!\"&\"b1\")", n(23.0)),
        (
            "VLOOKUP($A8,INDIRECT(\"'\"&$A$4&\"'!\"&\"a1:k110\"),B$2,FALSE)",
            n(23.0),
        ),
        ("INDIRECT(\"'(england)'!d6\")", n(29.0)),
        ("INDIRECT(\"'(\"&$A$3&\")'!d6\")", n(29.0)),
        ("INDIRECT(\"'O''Brien!'!a1\")", n(37.0)),
    ]);
}

#[test]
fn indirect_a1_new_structured_references_keep_cached_values() {
    for formula in [
        r#"INDIRECT("Table1[@[Unit Price]]")"#,
        r#"INDIRECT("Table1['#n]")"#,
    ] {
        let mut wb = workbook();
        let at = CellRef::parse_a1("D4").unwrap();
        wb.sheets[0].set_cell(
            at,
            Cell {
                formula: Some(formula.into()),
                value: n(42.0),
                ..Cell::default()
            },
        );
        let (graph, _) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
        assert_eq!(wb.value(SheetId(0), at), n(42.0), "{formula}");
        assert_eq!(
            graph.unevaluated_cells(),
            vec![xlsx_calc::UnevaluatedCell {
                sheet: SheetId(0),
                cell: at,
                functions: vec![],
            }],
            "{formula}"
        );
    }
}

#[test]
fn spaced_a1_cell_ranges_recalculate_to_release_errors() {
    let mut wb = workbook();
    let mut data = Sheet::new("Sheet1");
    set(&mut data, "A1", n(10.0));
    set(&mut data, "B2", n(20.0));
    wb.sheets.push(data);
    set(&mut wb.sheets[0], "A1", n(10.0));
    let cases = [
        ("D4", "SUM(Sheet1!A1 : B2)"),
        ("D5", r#"INDIRECT("A1 : B2")"#),
    ];
    for (address, formula) in cases {
        wb.sheets[0].set_cell(
            CellRef::parse_a1(address).unwrap(),
            Cell {
                formula: Some(formula.into()),
                value: n(42.0),
                ..Cell::default()
            },
        );
    }
    let (graph, _) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
    for (address, formula) in cases {
        assert_eq!(
            wb.value(SheetId(0), CellRef::parse_a1(address).unwrap()),
            CellValue::Error {
                value: ErrorValue::Ref,
            },
            "{formula}"
        );
    }
    assert!(graph.unevaluated_cells().is_empty());
}

#[test]
fn indirect_a1_bracketed_workbook_path_non_references_recalculate_to_ref() {
    let path = "D:/Books/[Data.xlsx]Sheet1";
    for target in ["#REF!", "A1+1"] {
        let mut wb = workbook();
        let mut path_named_sheet = Sheet::new(path);
        set(&mut path_named_sheet, "A1", n(73.0));
        wb.sheets.push(path_named_sheet);
        let at = CellRef::parse_a1("D4").unwrap();
        let formula = format!(r#"INDIRECT("'{path}'!{target}")"#);
        wb.sheets[0].set_cell(
            at,
            Cell {
                formula: Some(formula.clone()),
                value: n(42.0),
                ..Cell::default()
            },
        );
        let (graph, _) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
        assert_eq!(
            wb.value(SheetId(0), at),
            CellValue::Error {
                value: ErrorValue::Ref,
            },
            "{formula}"
        );
        assert!(graph.unevaluated_cells().is_empty(), "{formula}");
    }
}

#[test]
fn indirect_r1c1_absolute_and_relative_recalculate() {
    check(&[
        ("INDIRECT(Calculation_HIDE!S8,FALSE)", n(31.0)),
        ("INDIRECT(\"R5C2\",FALSE)", n(11.0)),
        ("INDIRECT(\"calculation_hide!r5c2\",0)", n(11.0)),
        ("INDIRECT(\"'main DATA'!r12c3\",FALSE)", n(31.0)),
        ("INDIRECT(\"'O''Brien!'!R1C1\",FALSE)", n(37.0)),
        ("INDIRECT(\"R[-1]C[2]\",FALSE)", n(13.0)),
        ("INDIRECT(\"R[2]C\",FALSE)", n(17.0)),
        ("INDIRECT(\"R5C[-1]\",FALSE)", n(11.0)),
        ("SUM(INDIRECT(\"R5C2:R6C3\",FALSE))", n(28.0)),
        ("SUM(INDIRECT(\"R6C3:R5C2\",FALSE))", n(28.0)),
        ("COLUMN(INDIRECT(\"RC\",FALSE))", n(3.0)),
        ("ROW(INDIRECT(\"RC\",FALSE))", n(4.0)),
        ("COLUMN(INDIRECT(\"R[-1]C[+2]\",FALSE))", n(5.0)),
    ]);
}

#[test]
fn indirect_r1c1_bracketed_workbook_paths_keep_the_cache_for_valid_references() {
    let ref_error = CellValue::Error {
        value: ErrorValue::Ref,
    };
    for (reference, expected) in [
        ("'D:/Books/[Data.xlsx]Sheet1'!R1C1", n(42.0)),
        (r"'C:\Books\[1]Sheet1'!R1C1", n(42.0)),
        ("'D:/Books/[Data.xlsx]Sheet1'!R0C1", ref_error.clone()),
        ("'D:/Books/[Data.xlsx]Sheet1' !R1C1", ref_error.clone()),
        ("'D:/Books/Data.xlsx'!R1C1", ref_error),
    ] {
        let mut wb = workbook();
        let at = CellRef::parse_a1("D4").unwrap();
        wb.sheets[0].set_cell(
            at,
            Cell {
                formula: Some(format!(r#"INDIRECT("{reference}",FALSE)"#)),
                value: n(42.0),
                ..Cell::default()
            },
        );
        let (graph, _) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
        assert_eq!(wb.value(SheetId(0), at), expected, "{reference}");
        assert_eq!(
            graph.unevaluated_cells().is_empty(),
            matches!(expected, CellValue::Error { .. }),
            "{reference}"
        );
    }
}

#[test]
fn indirect_r1c1_whole_rows_and_columns_recalculate() {
    check(&[
        ("SUM(INDIRECT(\"R5\",FALSE))", n(11.0)),
        ("SUM(INDIRECT(\"R5:R6\",FALSE))", n(28.0)),
        ("SUM(INDIRECT(\"R6:R5\",FALSE))", n(28.0)),
        ("SUM(INDIRECT(\"C5\",FALSE))", n(13.0)),
        ("SUM(INDIRECT(\"C5:C19\",FALSE))", n(13.0)),
        ("SUM(INDIRECT(\"C19:C5\",FALSE))", n(13.0)),
        ("SUM(INDIRECT(\"'main DATA'!C3\",FALSE))", n(31.0)),
        ("SUM(INDIRECT(\"'main DATA'!R12\",FALSE))", n(31.0)),
        ("ROWS(INDIRECT(\"C3:C5\",FALSE))", n(1_048_576.0)),
        ("COLUMNS(INDIRECT(\"R2:R5\",FALSE))", n(16_384.0)),
    ]);
}

#[test]
fn indirect_rejects_invalid_or_missing_references() {
    for reference in [
        "",
        "R0",
        "C0",
        "R0C1",
        "R1C0",
        "R1048577C1",
        "R1C16385",
        "R[-4]C",
        "R[1048576]C",
        "RC[16384]",
        "RC[-3]",
        "R[]C",
        "R[1C2",
        "R1C1junk",
        "R1C1:R2",
        "R2:C3",
        "A1",
        "SUM(R1C1)",
        "R1C1+1",
        "R1C1:R2C2:R3C3",
        "R-1C2",
        "R1.5C2",
        "R1C1!R1C1",
        "'missing'!R1C1",
        "R9999999999999999999999999C1",
        "R[9223372036854775807]C",
    ] {
        check(&[(
            &format!("INDIRECT(\"{reference}\",FALSE)"),
            CellValue::Error {
                value: ErrorValue::Ref,
            },
        )]);
    }
    check(&[
        (
            "INDIRECT(\"'missing'!a1\")",
            CellValue::Error {
                value: ErrorValue::Ref,
            },
        ),
        (
            "'missing'!a1",
            CellValue::Error {
                value: ErrorValue::Ref,
            },
        ),
        (
            "INDIRECT(\"R5C2\",1/0)",
            CellValue::Error {
                value: ErrorValue::Div0,
            },
        ),
    ]);
}

#[test]
fn sheet_qualified_ref_errors_recalculate_and_propagate() {
    for formula in [
        "'Monthly table (TIC)'!#REF!",
        "'S'!#REF!+1",
        "SUM('S'!#REF!)",
        "Calculation_HIDE!#REF!",
        "SUM(Calculation_HIDE!#REF!)",
        "#REF!",
        "#REF!+1",
        "SUM(#REF!)",
    ] {
        check(&[(
            formula,
            CellValue::Error {
                value: ErrorValue::Ref,
            },
        )]);
    }
}
