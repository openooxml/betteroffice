use xlsx_calc::{rebuild_and_recalc_all, recalc_after};
use xlsx_model::styles::Xf;
use xlsx_model::{Cell, CellRange, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn recalc(
    text: &str,
    formats: &[&str],
    builtin: Option<u16>,
    formulas: &[String],
) -> Vec<CellValue> {
    let mut wb = Workbook::default();
    for (index, code) in formats.iter().enumerate() {
        wb.styles
            .num_fmts
            .push((164 + index as u16, (*code).into()));
    }
    if let Some(id) = builtin {
        wb.styles.cell_xfs.push(Xf {
            num_fmt_id: Some(id),
            ..Xf::default()
        });
    }
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Text { value: text.into() },
            style: builtin.map(|_| wb.styles.cell_xfs.len() as u32 - 1),
            ..Cell::default()
        },
    );
    for (index, formula) in formulas.iter().enumerate() {
        sheet.set_cell(
            CellRef::new(index as u32, 1),
            Cell {
                formula: Some(formula.clone()),
                ..Cell::default()
            },
        );
    }
    wb.sheets.push(sheet);
    rebuild_and_recalc_all(&mut wb, None);
    (0..formulas.len())
        .map(|index| {
            wb.sheets[0]
                .cell(CellRef::new(index as u32, 1))
                .unwrap()
                .value
                .clone()
        })
        .collect()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn value_error() -> CellValue {
    CellValue::Error {
        value: ErrorValue::Value,
    }
}

#[test]
fn currency_text_coerces_in_operators_and_functions_with_format_evidence() {
    for (text, expected) in [
        ("£0.00", 0.0),
        ("£12", 12.0),
        ("£ 12", 12.0),
        ("-£12", -12.0),
        ("£-12", -12.0),
        ("+£12", 12.0),
        ("£+12", 12.0),
        ("12 €", 12.0),
        ("12€", 12.0),
        ("£1,234,567.50", 1_234_567.5),
        ("(£1,234.00)", -1234.0),
    ] {
        let mut formulas = [
            "A1*1",
            "A1+1",
            "A1-1",
            "A1/2",
            "-A1",
            "VALUE(A1)",
            "SUM(A1)",
            "SUM(A1:A1)",
            "SUM(A1*1)",
            "SUM(A1&\"\")",
        ]
        .map(String::from)
        .to_vec();
        formulas.extend([format!("\"{text}\"*1"), format!("VALUE(\"{text}\")")]);
        let actual = recalc(text, &["\"£\"#,##0.00", "[$€-407]0.00"], None, &formulas);
        let expected = [
            expected,
            expected + 1.0,
            expected - 1.0,
            expected / 2.0,
            -expected,
            expected,
            0.0,
            0.0,
            expected,
            expected,
            expected,
            expected,
        ]
        .map(number);
        assert_eq!(actual, expected, "{text}");
    }
}

#[test]
fn currency_text_rejects_malformed_or_unevidenced_symbols() {
    let formulas = ["A1*1", "A1+1", "A1-1", "A1/2", "-A1", "VALUE(A1)"].map(String::from);
    for text in ["£", "£1,23", "$$1", "£1£", "12£34", "£1,234,56", "1,234"] {
        assert_eq!(
            recalc(text, &["[$£-809]0.00", "\"$\"0.00"], None, &formulas),
            vec![value_error(); formulas.len()],
            "{text}"
        );
    }
    assert_eq!(
        recalc("$12", &["[$£-809]0.00"], None, &formulas),
        vec![value_error(); formulas.len()]
    );
    assert_eq!(
        recalc("£0.00", &[], None, &formulas),
        vec![value_error(); formulas.len()]
    );
}

#[test]
fn currency_evidence_includes_used_builtin_formats() {
    let formulas = [String::from("A1*1"), String::from("VALUE(A1)")];
    assert_eq!(
        recalc("$12", &[], Some(7), &formulas),
        vec![number(12.0); 2]
    );
    assert_eq!(recalc("$12", &[], None, &formulas), vec![value_error(); 2]);
}

#[test]
fn unused_builtin_currency_format_does_not_supply_evidence() {
    let mut wb = Workbook::default();
    wb.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            num_fmt_id: Some(7),
            ..Xf::default()
        },
    ];
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            formula: Some("VALUE(\"$12\")".into()),
            ..Cell::default()
        },
    );
    wb.sheets.push(sheet);
    rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(
        wb.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
        value_error()
    );
}

#[test]
fn day_first_workbook_format_preserves_main_date_text_behavior() {
    let formulas = [
        "A1+1",
        "-A1",
        "A1%",
        "VALUE(A1)",
        "DATEVALUE(A1)",
        "YEAR(A1)",
        "MONTH(A1)",
        "DAY(A1)",
    ]
    .map(String::from);
    for text in ["19/11/2021", "01/02/2021"] {
        let actual = recalc(text, &["dd/mm/yyyy"], None, &formulas);
        assert_eq!(actual, recalc(text, &[], None, &formulas), "{text}");
        assert_eq!(actual[..4], vec![value_error(); 4], "{text}");
        let expected = if text == "19/11/2021" {
            vec![value_error(); 4]
        } else {
            vec![number(44198.0), number(2021.0), number(1.0), number(2.0)]
        };
        assert_eq!(&actual[4..], expected.as_slice(), "{text}");
    }
}

fn currency_formula(formula: &str, evidence: bool, array: bool) -> CellValue {
    let mut wb = Workbook::default();
    if evidence {
        wb.styles.num_fmts.push((164, "£#,##0.00".into()));
    }
    let mut sheet = Sheet::new("Sheet1");
    for (row, value) in [
        number(1.0),
        number(2.0),
        CellValue::Text {
            value: "£12".into(),
        },
    ]
    .into_iter()
    .enumerate()
    {
        sheet.set_cell(
            CellRef::new(row as u32, 0),
            Cell {
                value,
                ..Cell::default()
            },
        );
    }
    let anchor = CellRef::new(0, 2);
    sheet.set_cell(
        anchor,
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    if array {
        sheet.set_array_formula(anchor, CellRange::new(anchor, anchor));
    }
    wb.sheets.push(sheet);
    rebuild_and_recalc_all(&mut wb, None);
    wb.sheets[0].cell(anchor).unwrap().value.clone()
}

fn currency_cases(cases: &[(&str, CellValue)], array: bool) {
    for (formula, expected) in cases {
        let actual = currency_formula(formula, true, array);
        match (&actual, expected) {
            (CellValue::Number { value: actual }, CellValue::Number { value: expected }) => {
                assert!(
                    (actual - expected).abs() < 1e-9,
                    "{formula}: {actual} != {expected}"
                );
            }
            _ => assert_eq!(&actual, expected, "{formula}"),
        }
        assert_eq!(
            currency_formula(formula, false, array),
            value_error(),
            "{formula}"
        );
    }
}

#[test]
fn array_aggregates_coerce_direct_currency_text() {
    currency_cases(&[("SUM(A1:A2,\"£12\")", number(15.0))], false);
    currency_cases(
        &[
            ("SUM(A1:A2*1,\"£12\")", number(15.0)),
            ("AVERAGE(A1:A2*1,\"£12\")", number(5.0)),
            ("MIN(A1:A2*1,\"£12\")", number(1.0)),
            ("MAX(A1:A2*1,\"£12\")", number(12.0)),
            ("PRODUCT(A1:A2*1,\"£12\")", number(24.0)),
        ],
        true,
    );
    for evidence in [false, true] {
        assert_eq!(currency_formula("SUM(A1:A3)", evidence, true), number(3.0));
        assert_eq!(
            currency_formula("SUM(A1:A3*1)", evidence, true),
            if evidence {
                number(15.0)
            } else {
                value_error()
            }
        );
    }
}

#[test]
fn sumproduct_coerces_scalar_currency_text() {
    currency_cases(&[("SUMPRODUCT(\"£12\",2)", number(24.0))], false);
    for evidence in [false, true] {
        assert_eq!(
            currency_formula("SUMPRODUCT(A1:A3)", evidence, false),
            number(3.0)
        );
    }
}

#[test]
fn financial_optional_arguments_coerce_currency_text() {
    currency_cases(
        &[
            ("FV(0,1,-10,\"£12\")", number(-2.0)),
            ("PV(0,1,-10,\"£12\")", number(-2.0)),
            ("PMT(0,1,12,\"£2\")", number(-14.0)),
            ("NPER(0,-10,10,\"£20\")", number(3.0)),
            ("FV(0,1,-10,0,\"£1\")", number(10.0)),
            ("PV(0,1,-10,0,\"£1\")", number(10.0)),
            ("PMT(0,1,10,0,\"£1\")", number(-10.0)),
            ("NPER(0,-10,10,0,\"£1\")", number(1.0)),
        ],
        false,
    );
}

#[test]
fn parity_coerces_currency_text() {
    currency_cases(
        &[
            ("ISEVEN(\"£12\")", CellValue::Bool { value: true }),
            ("ISODD(\"£12\")", CellValue::Bool { value: false }),
        ],
        false,
    );
}

#[test]
fn count_coerces_only_direct_currency_text() {
    for array in [false, true] {
        assert_eq!(currency_formula("COUNT(\"£12\")", true, array), number(1.0));
        assert_eq!(
            currency_formula("COUNT(\"£12\")", false, array),
            number(0.0)
        );
        for evidence in [false, true] {
            assert_eq!(
                currency_formula("COUNT(A1:A3)", evidence, array),
                number(2.0)
            );
            assert_eq!(currency_formula("COUNT(A3)", evidence, array), number(0.0));
        }
    }
}

#[test]
fn array_dimensions_and_indices_coerce_currency_text() {
    currency_cases(
        &[
            ("SUM(SEQUENCE(\"£2\"))", number(3.0)),
            ("SUM(CHOOSECOLS({1,2},\"£2\"))", number(2.0)),
            ("SUM(CHOOSEROWS({1;2},\"£2\"))", number(2.0)),
            ("SUM(EXPAND({1},\"£2\",1,0))", number(1.0)),
            ("SUM(EXPAND({1},1,\"£2\",0))", number(1.0)),
            ("SUM(TAKE({1;2},\"£1\"))", number(1.0)),
            ("SUM(DROP({1;2},\"£1\"))", number(2.0)),
            ("SUM(INDEX(VSTACK(1,2),{\"£2\"}))", number(2.0)),
        ],
        true,
    );
}

#[test]
fn regression_and_matrix_inputs_coerce_currency_text() {
    currency_cases(
        &[
            ("SUM(LINEST({\"£2\";4;6},{1;2;3}))", number(2.0)),
            ("SUM(LINEST({2;4;6},{\"£1\";2;3}))", number(2.0)),
            ("SUM(TREND({\"£2\";4;6},{1;2;3},{4}))", number(8.0)),
            ("SUM(TREND({2;4;6},{\"£1\";2;3},{4}))", number(8.0)),
            ("SUM(TREND({2;4;6},{1;2;3},{\"£4\"}))", number(8.0)),
            (
                "SUM(TREND({1;2;3;4},{1,0;0,1;1,1;2,1},{\"£1\",2}))",
                number(5.0),
            ),
            ("SUM(MMULT({\"£2\"},{3}))", number(6.0)),
            ("SUM(MMULT({2},{\"£3\"}))", number(6.0)),
        ],
        true,
    );
}

#[test]
fn text_function_numeric_options_coerce_currency_text() {
    currency_cases(
        &[
            ("NUMBERVALUE(\"£12\")", number(12.0)),
            ("NUMBERVALUE(\"£12%\")", number(0.12)),
            (
                "TEXTJOIN(\",\",\"£1\",\"a\",\"b\")",
                CellValue::Text {
                    value: "a,b".into(),
                },
            ),
            (
                "FIXED(1234,0,\"£1\")",
                CellValue::Text {
                    value: "1234".into(),
                },
            ),
        ],
        false,
    );
}

#[test]
fn criteria_coerce_currency_targets_and_values() {
    for (formula, expected, unevidenced) in [
        ("COUNTIF(A1:A2,\">£1\")", 1.0, 0.0),
        ("COUNTIF(A3,12)", 1.0, 0.0),
        ("COUNTIF(A3,\">1\")", 1.0, 0.0),
        ("COUNTIF(A3,\"£12\")", 1.0, 1.0),
        ("COUNTIF(A3,\"£*\")", 1.0, 1.0),
    ] {
        assert_eq!(
            currency_formula(formula, true, false),
            number(expected),
            "{formula}"
        );
        assert_eq!(
            currency_formula(formula, false, false),
            number(unevidenced),
            "{formula}"
        );
    }
}

#[test]
fn holiday_arguments_coerce_currency_text() {
    for (text, plain) in [
        ("WORKDAY(1,1,\"£2\")", "WORKDAY(1,1,2)"),
        ("WORKDAY.INTL(1,1,1,\"£2\")", "WORKDAY.INTL(1,1,1,2)"),
        ("NETWORKDAYS(1,5,\"£2\")", "NETWORKDAYS(1,5,2)"),
        (
            "NETWORKDAYS.INTL(1,5,1,\"£2\")",
            "NETWORKDAYS.INTL(1,5,1,2)",
        ),
    ] {
        let expected = currency_formula(plain, true, false);
        assert!(matches!(expected, CellValue::Number { .. }), "{plain}");
        assert_eq!(currency_formula(text, true, false), expected, "{text}");
        assert_eq!(
            currency_formula(text, false, false),
            value_error(),
            "{text}"
        );
    }
}

#[test]
fn currency_evidence_follows_style_edits() {
    let (a1, b1, c1) = (CellRef::new(0, 0), CellRef::new(0, 1), CellRef::new(0, 2));
    let mut wb = Workbook::default();
    wb.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            num_fmt_id: Some(7),
            ..Xf::default()
        },
    ];
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        a1,
        Cell {
            value: CellValue::Text {
                value: "$12".into(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        b1,
        Cell {
            formula: Some("A1*1".into()),
            ..Cell::default()
        },
    );
    sheet.set_cell(
        c1,
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    wb.sheets.push(sheet);
    let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(wb.sheets[0].cell(b1).unwrap().value, value_error());

    // a cell takes the built-in currency format: the next recalc coerces
    wb.sheets[0].cell_mut(c1).unwrap().style = Some(1);
    recalc_after(&mut wb, &mut graph, &[(SheetId(0), c1)], None);
    assert_eq!(wb.sheets[0].cell(b1).unwrap().value, number(12.0));

    // the format is removed again: the next recalc stops coercing
    wb.sheets[0].cell_mut(c1).unwrap().style = None;
    recalc_after(&mut wb, &mut graph, &[(SheetId(0), c1)], None);
    assert_eq!(wb.sheets[0].cell(b1).unwrap().value, value_error());
}

#[test]
fn shadowed_column_currency_format_does_not_supply_evidence() {
    let mut wb = Workbook::default();
    wb.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            num_fmt_id: Some(7),
            ..Xf::default()
        },
    ];
    let mut sheet = Sheet::new("Sheet1");
    sheet.col_styles = vec![
        xlsx_model::ColStyle {
            first: 0,
            last: 5,
            xf: 1,
        },
        xlsx_model::ColStyle {
            first: 0,
            last: 5,
            xf: 0,
        },
    ];
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            formula: Some("\"$12\"+0".into()),
            ..Cell::default()
        },
    );
    wb.sheets.push(sheet);
    rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(
        wb.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
        value_error()
    );
    wb.sheets[0].col_styles[1].last = 4;
    rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(
        wb.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
        number(12.0)
    );
}

#[test]
fn currency_plus_percent_stays_text() {
    assert_eq!(currency_formula("\"12%\"+0", true, false), number(0.12));
    assert_eq!(currency_formula("\"£12%\"+0", true, false), value_error());
}

#[test]
fn standalone_contexts_share_workbook_currency_evidence() {
    let mut wb = Workbook::default();
    wb.styles.num_fmts.push((164, "£0.00".into()));
    wb.sheets.push(Sheet::new("Sheet1"));
    let expr = xlsx_calc::parse_formula("\"£12\"+1").unwrap();
    for ctx in [
        xlsx_calc::EvalContext::new(&wb, SheetId(0)),
        xlsx_calc::EvalContext::with_now(&wb, SheetId(0), 1.0),
    ] {
        assert_eq!(xlsx_calc::evaluate(&expr, &ctx), number(13.0));
    }
}
