//! RAND, RANK.AVG, FORECAST, the integer math bundle and FIXED against the
//! worked examples in Microsoft's function reference.

use xlsx_calc::{EvalContext, evaluate, evaluate_array, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}
fn t(v: &str) -> CellValue {
    CellValue::Text { value: v.into() }
}
fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn workbook(cells: &[(&str, CellValue)]) -> Workbook {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    for (a1, value) in cells {
        sheet.set_cell(
            CellRef::parse_a1(a1).unwrap(),
            Cell {
                value: value.clone(),
                ..Cell::default()
            },
        );
    }
    wb.sheets.push(sheet);
    wb
}

fn eval_in(wb: &Workbook, src: &str) -> CellValue {
    let expr = parse_formula(src).expect("parse");
    evaluate(&expr, &EvalContext::new(wb, SheetId(0)))
}

fn check_in(wb: &Workbook, cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval_in(wb, src), *want, "formula {src:?}");
    }
}

fn check(cases: &[(&str, CellValue)]) {
    check_in(&workbook(&[]), cases);
}

fn approx_in(wb: &Workbook, src: &str, want: f64) {
    match eval_in(wb, src) {
        CellValue::Number { value } => {
            assert!(
                (value - want).abs() <= 1e-6 * want.abs().max(1.0),
                "formula {src:?}: {value} != {want}"
            );
        }
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

#[test]
fn rand_draws_from_the_unit_interval() {
    let wb = workbook(&[]);
    for _ in 0..200 {
        match eval_in(&wb, "RAND()") {
            CellValue::Number { value } => assert!((0.0..1.0).contains(&value), "{value}"),
            other => panic!("RAND() gave {other:?}"),
        }
        match eval_in(&wb, "RAND()*100") {
            CellValue::Number { value } => assert!((0.0..100.0).contains(&value), "{value}"),
            other => panic!("RAND()*100 gave {other:?}"),
        }
    }
    check(&[
        ("ISNUMBER(RAND())", CellValue::Bool { value: true }),
        ("RAND()<1", CellValue::Bool { value: true }),
        ("RAND()>=0", CellValue::Bool { value: true }),
        ("RAND(1)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn rand_shares_the_seeded_stream_with_randbetween() {
    let wb = workbook(&[]);
    let draw = |src: &str| {
        let expr = parse_formula(src).unwrap();
        let mut ctx = EvalContext::new(&wb, SheetId(0));
        ctx.rand_seed = Some(42);
        evaluate(&expr, &ctx)
    };
    assert_eq!(draw("RAND()"), draw("RAND()"));
    let (CellValue::Number { value: unit }, CellValue::Number { value: int }) =
        (draw("RAND()"), draw("RANDBETWEEN(0,999)"))
    else {
        panic!("seeded draws are numbers");
    };
    assert_eq!((unit * 1000.0).floor(), int);
}

#[test]
fn rank_avg_averages_tied_ranks() {
    // Microsoft example: temperatures in B2:B8
    let wb = workbook(&[
        ("A1", t("Day")),
        ("B1", t("Temp (F)")),
        ("A2", n(40725.0)),
        ("B2", n(89.0)),
        ("A3", n(40726.0)),
        ("B3", n(88.0)),
        ("A4", n(40727.0)),
        ("B4", n(92.0)),
        ("A5", n(40728.0)),
        ("B5", n(101.0)),
        ("A6", n(40729.0)),
        ("B6", n(94.0)),
        ("A7", n(40730.0)),
        ("B7", n(97.0)),
        ("A8", n(40731.0)),
        ("B8", n(95.0)),
        ("D1", n(1.0)),
        ("D2", n(2.0)),
        ("D3", n(2.0)),
        ("D4", n(3.0)),
        ("D5", t("not numeric")),
        ("D6", n(2.0)),
    ]);
    check_in(
        &wb,
        &[
            ("RANK.AVG(94,B2:B8)", n(4.0)),
            ("_xlfn.RANK.AVG(94,B2:B8)", n(4.0)),
            ("RANK.AVG(101,B2:B8)", n(1.0)),
            ("RANK.AVG(88,B2:B8,1)", n(1.0)),
            ("RANK.AVG(2,D1:D6)", n(3.0)),
            ("RANK.AVG(2,D1:D6,1)", n(3.0)),
            ("RANK.AVG(1,D1:D6)", n(5.0)),
            ("RANK.AVG(3,D1:D6,1)", n(5.0)),
            ("RANK.AVG(2,D1:D4)", n(2.5)),
            ("RANK.AVG(5,D1:D6)", e(ErrorValue::NA)),
            ("RANK.AVG(2,{1,2,3})", e(ErrorValue::Value)),
            ("RANK.EQ(2,D1:D6)", n(2.0)),
        ],
    );
    let expr = parse_formula("RANK.AVG(D1:D4,D1:D4)").unwrap();
    let ranks = evaluate_array(&expr, &EvalContext::new(&wb, SheetId(0)));
    assert_eq!(
        ranks.into_array().into_values(),
        vec![n(4.0), n(2.5), n(2.5), n(1.0)]
    );
}

#[test]
fn forecast_extends_the_least_squares_line() {
    // Microsoft example: known_y in A2:A6, known_x in B2:B6
    let wb = workbook(&[
        ("A1", t("Known Y")),
        ("B1", t("Known X")),
        ("A2", n(6.0)),
        ("B2", n(20.0)),
        ("A3", n(7.0)),
        ("B3", n(28.0)),
        ("A4", n(9.0)),
        ("B4", n(31.0)),
        ("A5", n(15.0)),
        ("B5", n(38.0)),
        ("A6", n(21.0)),
        ("B6", n(40.0)),
        ("C1", n(1.0)),
        ("C2", n(1.0)),
    ]);
    approx_in(&wb, "FORECAST(30,A2:A6,B2:B6)", 10.607253);
    approx_in(&wb, "FORECAST.LINEAR(30,A2:A6,B2:B6)", 10.607253);
    approx_in(&wb, "_xlfn.FORECAST.LINEAR(30,A2:A6,B2:B6)", 10.607253);
    check_in(
        &wb,
        &[
            ("FORECAST(4,{1,2,3},{1,2,3})", n(4.0)),
            ("FORECAST(4,{2,4,6},{1,2,3})", n(8.0)),
            ("FORECAST.LINEAR(4,{3,6,9},{1,2,3})", n(12.0)),
            ("FORECAST(\"x\",A2:A6,B2:B6)", e(ErrorValue::Value)),
            ("FORECAST(30,A2:A6,B2:B5)", e(ErrorValue::NA)),
            ("FORECAST(30,A2:A3,C1:C2)", e(ErrorValue::Div0)),
            ("FORECAST(30,A2:A6)", e(ErrorValue::Value)),
        ],
    );
}

#[test]
fn integer_math_checks_signs_before_truncating_and_reads_blanks_as_zero() {
    // B1 is blank
    let wb = workbook(&[("B2", n(5.0))]);
    check_in(
        &wb,
        &[
            ("FACT(-0.5)", e(ErrorValue::Num)),
            ("FACT(-0)", n(1.0)),
            ("COMBIN(-0.5,0)", e(ErrorValue::Num)),
            ("COMBIN(5,-0.5)", e(ErrorValue::Num)),
            ("LCM(9007199254740991,2)", e(ErrorValue::Num)),
            ("LCM(100000000,100000001)", e(ErrorValue::Num)),
            ("GCD(9007199254740991,2)", n(1.0)),
            ("LCM(B1,5)", n(0.0)),
            ("LCM(B1:B2)", n(0.0)),
            ("LCM(,5)", n(0.0)),
            ("GCD(B1,5)", n(5.0)),
            ("GCD(B1:B2)", n(5.0)),
            ("GCD(,5)", n(5.0)),
            ("GCD(B1)", n(0.0)),
        ],
    );
}

#[test]
fn integer_math_follows_the_reference_examples() {
    let wb = workbook(&[("A1", n(24.0)), ("A2", n(36.0)), ("A25", n(8.0))]);
    check_in(
        &wb,
        &[
            ("GCD(5, 2)", n(1.0)),
            ("GCD(24, 36)", n(12.0)),
            ("GCD(7, 1)", n(1.0)),
            ("GCD(5, 0)", n(5.0)),
            ("GCD(A1:A2)", n(12.0)),
            ("GCD(24.9, 36.2)", n(12.0)),
            ("GCD(-24, 36)", e(ErrorValue::Num)),
            ("GCD(\"x\", 36)", e(ErrorValue::Value)),
            ("GCD(9007199254740992, 2)", e(ErrorValue::Num)),
            ("LCM(5, 2)", n(10.0)),
            ("LCM(24, 36)", n(72.0)),
            ("LCM(1,6)", n(6.0)),
            ("LCM(A1:A2, 5)", n(360.0)),
            ("LCM(5, 0)", n(0.0)),
            ("LCM(-5, 2)", e(ErrorValue::Num)),
            ("EVEN(1.5)", n(2.0)),
            ("EVEN(3)", n(4.0)),
            ("EVEN(2)", n(2.0)),
            ("EVEN(-1)", n(-2.0)),
            ("EVEN(0)", n(0.0)),
            ("EVEN(\"x\")", e(ErrorValue::Value)),
            ("ODD(1.5)", n(3.0)),
            ("ODD(3)", n(3.0)),
            ("ODD(2)", n(3.0)),
            ("ODD(-1)", n(-1.0)),
            ("ODD(-2)", n(-3.0)),
            ("ODD(0)", n(1.0)),
            ("FACT(5)", n(120.0)),
            ("FACT(1.9)", n(1.0)),
            ("FACT(0)", n(1.0)),
            ("FACT(-1)", e(ErrorValue::Num)),
            ("FACT(1)", n(1.0)),
            ("FACT(171)", e(ErrorValue::Num)),
            ("COMBIN(8,2)", n(28.0)),
            ("COMBIN(A25,2)", n(28.0)),
            ("COMBIN(A25:A25,2)", n(28.0)),
            ("COMBIN(8.9,2.9)", n(28.0)),
            ("COMBIN(8,0)", n(1.0)),
            ("COMBIN(8,8)", n(1.0)),
            ("COMBIN(2,3)", e(ErrorValue::Num)),
            ("COMBIN(-1,0)", e(ErrorValue::Num)),
            ("COMBIN(8,-1)", e(ErrorValue::Num)),
            ("COMBIN(\"x\",2)", e(ErrorValue::Value)),
        ],
    );
    let wb = workbook(&[]);
    approx_in(&wb, "FACT(170)", 7.257415615307994e306);
    approx_in(&wb, "COMBIN(60,30)", 1.1826458156486e17);
}

#[test]
fn fixed_rounds_and_formats_as_text() {
    // Microsoft example: data in A2:A4
    let wb = workbook(&[
        ("A1", t("Data")),
        ("A2", n(1234.567)),
        ("A3", n(-1234.567)),
        ("A4", n(44.332)),
    ]);
    check_in(
        &wb,
        &[
            ("FIXED(A2, 1)", t("1,234.6")),
            ("FIXED(A2, -1)", t("1,230")),
            ("FIXED(A3, -1, TRUE)", t("-1230")),
            ("FIXED(A4)", t("44.33")),
            ("FIXED(A2)", t("1,234.57")),
            ("FIXED(A3, 2)", t("-1,234.57")),
            ("FIXED(A2, 0)", t("1,235")),
            ("FIXED(A2, -4)", t("0")),
            ("FIXED(A2, -3)", t("1,000")),
            ("FIXED(1234567.891, 2, FALSE)", t("1,234,567.89")),
            ("FIXED(1234567.891, 2, TRUE)", t("1234567.89")),
            ("FIXED(2.675, 2)", t("2.68")),
            ("FIXED(0.5, 0)", t("1")),
            ("FIXED(-0.001, 2)", t("0.00")),
            ("FIXED(0.1, 20)", t("0.10000000000000000000")),
            ("FIXED(999.996, 2)", t("1,000.00")),
            ("FIXED(\"12.5\", 1)", t("12.5")),
            ("FIXED(1, 128)", e(ErrorValue::Value)),
            ("FIXED(\"x\")", e(ErrorValue::Value)),
            ("FIXED(1/0)", e(ErrorValue::Div0)),
        ],
    );
}

#[test]
fn combin_stays_exact_below_2_53_and_overflows_only_past_f64() {
    let wb = workbook(&[]);
    check(&[
        ("COMBIN(56,24)", n(4_355_031_703_297_275.0)),
        ("COMBIN(1030,515)", e(ErrorValue::Num)),
    ]);
    approx_in(&wb, "COMBIN(1028,514)", 7.156051054877897e307);
    approx_in(&wb, "COMBIN(1029,514)", 1.429820686498904e308);
    approx_in(&wb, "COMBIN(66,33)", 7.219428434016266e18);
}

#[test]
fn scalar_functions_answer_once_per_array_element() {
    let wb = workbook(&[]);
    let ctx = EvalContext::new(&wb, SheetId(0));
    let cases: &[(&str, Vec<CellValue>)] = &[
        ("EVEN({1;3})", vec![n(2.0), n(4.0)]),
        ("ODD({1.5;2})", vec![n(3.0), n(3.0)]),
        ("FACT({3;4})", vec![n(6.0), n(24.0)]),
        ("COMBIN({8;10},2)", vec![n(28.0), n(45.0)]),
        ("COMBIN(8,{2;3})", vec![n(28.0), n(56.0)]),
        ("FIXED({1234.5;2},1)", vec![t("1,234.5"), t("2.0")]),
        ("FIXED(1234.5,{0;1},TRUE)", vec![t("1235"), t("1234.5")]),
        ("FORECAST({4;5},{1,2,3},{1,2,3})", vec![n(4.0), n(5.0)]),
        (
            "FORECAST.LINEAR({4;5},{2,4,6},{1,2,3})",
            vec![n(8.0), n(10.0)],
        ),
        ("GCD({24;36})", vec![n(12.0)]),
    ];
    for (src, want) in cases {
        let expr = parse_formula(src).expect("parse");
        let got = evaluate_array(&expr, &ctx).into_array().into_values();
        assert_eq!(&got, want, "formula {src:?}");
    }
}
