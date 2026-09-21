//! table-driven coverage for the function library: each case parses a formula,
//! evaluates it against a shared in-memory workbook, and asserts the value.

use xlsx_calc::{EvalContext, evaluate, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}
fn t(v: &str) -> CellValue {
    CellValue::Text { value: v.into() }
}
fn b(v: bool) -> CellValue {
    CellValue::Bool { value: v }
}
fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

/// fixture: A1:A5 = 10..50, B1:B5 = fruit names, C1:C5 = 1..5, E1:F4 vertical
/// and H1:K2 horizontal lookup tables.
fn fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    let put = |s: &mut Sheet, a1: &str, v: CellValue| {
        s.set_cell(
            CellRef::parse_a1(a1).unwrap(),
            Cell {
                value: v,
                ..Cell::default()
            },
        );
    };
    for (i, v) in [10.0, 20.0, 30.0, 40.0, 50.0].iter().enumerate() {
        put(&mut s, &format!("A{}", i + 1), n(*v));
    }
    for (i, v) in ["apple", "banana", "apple", "cherry", "apple"]
        .iter()
        .enumerate()
    {
        put(&mut s, &format!("B{}", i + 1), t(v));
    }
    for (i, v) in [1.0, 2.0, 3.0, 4.0, 5.0].iter().enumerate() {
        put(&mut s, &format!("C{}", i + 1), n(*v));
    }
    let names = ["one", "two", "three", "four"];
    for (i, name) in names.iter().enumerate() {
        put(&mut s, &format!("E{}", i + 1), n(i as f64 + 1.0));
        put(&mut s, &format!("F{}", i + 1), t(name));
    }
    for (i, col) in ["H", "I", "J", "K"].iter().enumerate() {
        put(&mut s, &format!("{col}1"), n(i as f64 + 1.0));
        put(&mut s, &format!("{col}2"), t(["a", "b", "c", "d"][i]));
    }
    wb.sheets.push(s);
    wb
}

fn eval(src: &str) -> CellValue {
    let wb = fixture();
    let expr = parse_formula(src).expect("parse");
    let ctx = EvalContext::new(&wb, SheetId(0));
    evaluate(&expr, &ctx)
}

#[test]
fn vlookup_accepts_whole_columns() {
    for columns in [
        "E:F",
        "$E:$F",
        "E:$F",
        "$E:F",
        "Sheet1!E:F",
        "'Sheet1'!$E:$F",
    ] {
        assert_eq!(eval(&format!("VLOOKUP(2,{columns},2,FALSE)")), t("two"));
    }
    check(&[
        ("VLOOKUP(2.5,E:F,2)", t("two")),
        ("VLOOKUP(9,E:F,2,FALSE)", e(ErrorValue::NA)),
        ("VLOOKUP(2,E:F,3,FALSE)", e(ErrorValue::Ref)),
        ("VLOOKUP(2,E:F,0,FALSE)", e(ErrorValue::Value)),
        ("SUM(E:E)", n(10.0)),
        ("MATCH(2,E:E,0)", n(2.0)),
    ]);
}

/// like `eval` but with an injected clock (2020-01-01 12:00) for TODAY/NOW.
fn eval_now(src: &str) -> CellValue {
    let wb = fixture();
    let expr = parse_formula(src).expect("parse");
    let ctx = EvalContext::with_now(&wb, SheetId(0), 43_831.5);
    evaluate(&expr, &ctx)
}

fn check(cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "formula {src:?}");
    }
}

fn approx(src: &str, want: f64) {
    match eval(src) {
        CellValue::Number { value } => {
            assert!(
                (value - want).abs() < 1e-9,
                "formula {src:?}: {value} != {want}"
            );
        }
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

#[test]
fn math_functions() {
    check(&[
        ("SUMIF(A1:A5, \">=30\")", n(120.0)),
        ("SUMIF(B1:B5, \"apple\", C1:C5)", n(9.0)),
        ("SUMIFS(C1:C5, B1:B5, \"apple\", A1:A5, \">=30\")", n(8.0)),
        ("SUMPRODUCT(A1:A3, C1:C3)", n(140.0)),
        ("PRODUCT(1, 2, 3, 4)", n(24.0)),
        ("ROUNDUP(2.1, 0)", n(3.0)),
        ("ROUNDDOWN(2.9, 0)", n(2.0)),
        ("INDIRECT(\"A1\")", n(10.0)),
        ("INDIRECT(\"A\" & \"2\")", n(20.0)),
        ("SUM(INDIRECT(\"A1:A3\"))", n(60.0)),
        ("INDIRECT(\"B1\")", t("apple")),
        ("INDIRECT(\"nonsense!!\")", e(ErrorValue::Ref)),
        ("INDIRECT(\"SUM(A1:A3)\")", e(ErrorValue::Ref)),
        ("INDIRECT(\"A1\", FALSE)", e(ErrorValue::Ref)),
        ("INDIRECT()", e(ErrorValue::Value)),
        ("ROUND(SIN(PI()/2), 9)", n(1.0)),
        ("ROUND(COS(0), 9)", n(1.0)),
        ("ROUND(TAN(0), 9)", n(0.0)),
        ("ROUND(DEGREES(PI()), 9)", n(180.0)),
        ("ROUND(RADIANS(180) - PI(), 9)", n(0.0)),
        ("ASIN(2)", e(ErrorValue::Num)),
        ("ACOS(-2)", e(ErrorValue::Num)),
        ("ROUND(ATAN2(1, 1) * 4 - PI(), 9)", n(0.0)),
        ("ATAN2(0, 0)", e(ErrorValue::Div0)),
        ("ROUND(SLOPE(A1:A5, B1:B5), 9)", e(ErrorValue::NA)),
        ("ROUND(CORREL(A1:A5, C1:C5), 9)", n(1.0)),
        ("ROUND(SLOPE(A1:A5, C1:C5), 9)", n(10.0)),
        ("ROUND(INTERCEPT(A1:A5, C1:C5), 9)", n(0.0)),
        ("ROUND(COVARIANCE.P(A1:A5, C1:C5), 9)", n(20.0)),
        ("ROUND(COVARIANCE.S(A1:A5, C1:C5), 9)", n(25.0)),
        ("PERCENTILE.INC(A1:A5, 0)", n(10.0)),
        ("PERCENTILE.INC(A1:A5, 1)", n(50.0)),
        ("PERCENTILE.INC(A1:A5, 0.5)", n(30.0)),
        ("PERCENTILE.INC(A1:A5, 2)", e(ErrorValue::Num)),
        ("QUARTILE.INC(A1:A5, 0)", n(10.0)),
        ("QUARTILE.INC(A1:A5, 2)", n(30.0)),
        ("QUARTILE.INC(A1:A5, 4)", n(50.0)),
        ("QUARTILE.INC(A1:A5, 5)", e(ErrorValue::Num)),
        ("ISOWEEKNUM(46023)", n(1.0)),
        ("ISOWEEKNUM(45656)", n(1.0)),
        ("ISOWEEKNUM(44200)", n(1.0)),
        ("ISOWEEKNUM(44196)", n(53.0)),
        ("ISOWEEKNUM(46286)", n(39.0)),
        ("WEEKNUM(46023, 21)", n(1.0)),
        ("WEEKNUM(44196, 21)", n(53.0)),
        ("WEEKNUM(46023)", n(1.0)),
        ("WEEKNUM(46023, 99)", e(ErrorValue::Num)),
        ("TEXTAFTER(\"\u{130}a\", \"a\", 1, 1)", t("")),
        ("TEXTBEFORE(\"\u{130}a\", \"a\", 1, 1)", t("\u{130}")),
        ("TEXTBEFORE(\"stra\u{df}e-x\", \"-\")", t("stra\u{df}e")),
        ("TEXTBEFORE(\"a-b-c\", \"-\")", t("a")),
        ("TEXTAFTER(\"a-b-c\", \"-\")", t("b-c")),
        ("TEXTBEFORE(\"a-b-c\", \"-\", 2)", t("a-b")),
        ("TEXTAFTER(\"a-b-c\", \"-\", 2)", t("c")),
        ("TEXTBEFORE(\"a-b-c\", \"-\", -1)", t("a-b")),
        ("TEXTAFTER(\"a-b-c\", \"-\", -2)", t("b-c")),
        ("TEXTBEFORE(\"a-b\", \"X\")", e(ErrorValue::NA)),
        ("TEXTBEFORE(\"a-b\", \"X\", 1, 0, \"none\")", t("none")),
        ("TEXTBEFORE(\"aXb\", \"x\", 1, 1)", t("a")),
        ("TEXTAFTER(\"a-b-c\", \"-\", 9)", e(ErrorValue::NA)),
        ("AGGREGATE(14, 6, A1:A5, 2)", n(40.0)),
        ("AGGREGATE(12, 6, A1:A5)", n(30.0)),
        ("AGGREGATE(7, 6, A1:A5)", n(15.811388300841896)),
        ("XMATCH(30, A1:A5, 2)", e(ErrorValue::Value)),
        ("XMATCH(30, A1:A5, 0, 3)", e(ErrorValue::Value)),
        ("INDIRECT(\"A1\", 1/0)", e(ErrorValue::Div0)),
        ("SERIESSUM(2, 0, 1, B1:B5)", e(ErrorValue::Value)),
        ("SUBTOTAL(9, A1:A5)", n(150.0)),
        ("SUBTOTAL(109, A1:A5)", n(150.0)),
        ("SUBTOTAL(1, A1:A5)", n(30.0)),
        ("SUBTOTAL(4, A1:A5)", n(50.0)),
        ("SUBTOTAL(2, A1:A5)", n(5.0)),
        ("SUBTOTAL(99, A1:A5)", e(ErrorValue::Value)),
        ("SUBTOTAL(9)", e(ErrorValue::Value)),
        ("AGGREGATE(9, 0, A1:A5)", n(150.0)),
        ("AGGREGATE(1, 0, A1:A5)", n(30.0)),
        ("AGGREGATE(14, 0, A1:A5, 2)", n(40.0)),
        ("AGGREGATE(15, 0, A1:A5, 2)", n(20.0)),
        ("AGGREGATE(9, 8, A1:A5)", e(ErrorValue::Value)),
        ("AGGREGATE(99, 0, A1:A5)", e(ErrorValue::Value)),
        ("XMATCH(30, A1:A5)", n(3.0)),
        ("XMATCH(30, A1:A5, 0)", n(3.0)),
        ("XMATCH(35, A1:A5, -1)", n(3.0)),
        ("XMATCH(35, A1:A5, 1)", n(4.0)),
        ("XMATCH(35, A1:A5, 0)", e(ErrorValue::NA)),
        ("XMATCH(30, A1:A5, 0, -1)", n(3.0)),
        ("QUOTIENT(5, 2)", n(2.0)),
        ("QUOTIENT(-5, 2)", n(-2.0)),
        ("QUOTIENT(5, -2)", n(-2.0)),
        ("QUOTIENT(4.9, 2)", n(2.0)),
        ("QUOTIENT(1, 0)", e(ErrorValue::Div0)),
        ("QUOTIENT(1)", e(ErrorValue::Value)),
        ("SERIESSUM(2, 0, 1, C1:C5)", n(129.0)),
        ("SERIESSUM(2, 1, 2, C1:C3)", n(114.0)),
        ("SERIESSUM(2, 0, 1)", e(ErrorValue::Value)),
        ("ISEVEN(4)", b(true)),
        ("ISEVEN(3.9)", b(false)),
        ("ISEVEN(-4)", b(true)),
        ("ISEVEN(-3)", b(false)),
        ("ISODD(3)", b(true)),
        ("ISODD(-3)", b(true)),
        ("ISODD(4)", b(false)),
        ("ISODD(0)", b(false)),
        ("MROUND(10, 3)", n(9.0)),
        ("MROUND(-2.5, -1)", n(-3.0)),
        ("INT(-2.5)", n(-3.0)),
        ("TRUNC(-2.7)", n(-2.0)),
        ("TRUNC(1.98765, 2)", n(1.98)),
        ("MOD(-3, 2)", n(1.0)),
        ("POWER(2, 10)", n(1024.0)),
        ("SQRT(16)", n(4.0)),
        ("SQRT(-1)", e(ErrorValue::Num)),
        ("LOG(8, 2)", n(3.0)),
        ("LOG10(1000)", n(3.0)),
        ("SIGN(-5)", n(-1.0)),
        ("CEILING(2.1, 1)", n(3.0)),
        ("FLOOR(2.9, 1)", n(2.0)),
        ("CEILING(-2.5, -1)", n(-3.0)),
        ("ABS(-7)", n(7.0)),
        ("TANH(0)", n(0.0)),
        ("TANH(\"abc\")", e(ErrorValue::Value)),
        ("TANH(1, 2)", e(ErrorValue::Value)),
    ]);
    approx("PI()", std::f64::consts::PI);
    approx("LN(EXP(1))", 1.0);
    approx("EXP(0)", 1.0);
    approx("TANH(1)", 0.761_594_155_955_764_9);
    approx("TANH(-2.5)", -0.986_614_298_151_430_3);
    approx("TANH(20)", 1.0);
    approx("TANH(TRUE)", 0.761_594_155_955_764_9);
}

#[test]
fn stats_functions() {
    check(&[
        ("MEDIAN(1, 2, 3, 4)", n(2.5)),
        ("MEDIAN(A1:A5)", n(30.0)),
        ("MODE(1, 2, 2, 3)", n(2.0)),
        ("MODE(1, 2, 3)", e(ErrorValue::NA)),
        ("VARP(2, 4, 4, 4, 5, 5, 7, 9)", n(4.0)),
        ("STDEVP(2, 4, 4, 4, 5, 5, 7, 9)", n(2.0)),
        ("VAR(1, 2, 3, 4, 5)", n(2.5)),
        ("LARGE(A1:A5, 1)", n(50.0)),
        ("LARGE(A1:A5, 2)", n(40.0)),
        ("SMALL(A1:A5, 2)", n(20.0)),
        ("RANK(30, A1:A5)", n(3.0)),
        ("RANK(30, A1:A5, 1)", n(3.0)),
        ("COUNTIF(B1:B5, \"apple\")", n(3.0)),
        ("COUNTIF(B1:B5, \"a*\")", n(3.0)),
        ("COUNTIF(B1:B5, \"<>apple\")", n(2.0)),
        ("COUNTIFS(B1:B5, \"apple\", A1:A5, \">=30\")", n(2.0)),
        ("COUNTBLANK(A1:A6)", n(1.0)),
        ("AVERAGEIF(A1:A5, \">=30\")", n(40.0)),
        ("AVERAGEIFS(C1:C5, B1:B5, \"apple\")", n(3.0)),
    ]);
}

#[test]
fn text_functions() {
    check(&[
        ("LEFT(\"hello\", 2)", t("he")),
        ("LEFT(\"hello\")", t("h")),
        ("RIGHT(\"hello\", 2)", t("lo")),
        ("MID(\"hello\", 2, 3)", t("ell")),
        ("FIND(\"l\", \"hello\")", n(3.0)),
        ("FIND(\"L\", \"hello\")", e(ErrorValue::Value)),
        ("SEARCH(\"L\", \"hello\")", n(3.0)),
        ("SUBSTITUTE(\"a-b-c\", \"-\", \"+\")", t("a+b+c")),
        ("SUBSTITUTE(\"a-b-c\", \"-\", \"+\", 2)", t("a-b+c")),
        ("REPLACE(\"abcdef\", 2, 3, \"XY\")", t("aXYef")),
        ("REPT(\"ab\", 3)", t("ababab")),
        ("EXACT(\"a\", \"a\")", b(true)),
        ("EXACT(\"a\", \"A\")", b(false)),
        ("PROPER(\"hello world\")", t("Hello World")),
        ("CLEAN(CHAR(7) & \"a\")", t("a")),
        ("CHAR(65)", t("A")),
        ("CODE(\"A\")", n(65.0)),
        ("VALUE(\"12.5\")", n(12.5)),
        ("VALUE(\"50%\")", n(0.5)),
        ("NUMBERVALUE(\"1,234.5\")", n(1234.5)),
        ("T(\"hi\")", t("hi")),
        ("T(5)", t("")),
        ("TEXTJOIN(\"-\", TRUE, \"a\", \"\", \"b\")", t("a-b")),
        ("TEXTJOIN(\"-\", FALSE, \"a\", \"\", \"b\")", t("a--b")),
        ("TEXT(1234.5, \"#,##0.00\")", t("1,234.50")),
        ("TEXT(0.5, \"0%\")", t("50%")),
        ("TEXT(2.5, \"0.00\")", t("2.50")),
        ("TEXT(0.1234, \"0.0%\")", t("12.3%")),
        ("TEXT(-5, \"0.00;(0.00)\")", t("(5.00)")),
        ("TEXT(12345, \"0.00E+00\")", t("1.23E+04")),
        ("TEXT(43831, \"m/d/yyyy\")", t("1/1/2020")),
        ("TEXT(43831, \"mmmm d, yyyy\")", t("January 1, 2020")),
        ("TEXT(0.5, \"h:mm AM/PM\")", t("12:00 PM")),
        ("TEXT(5, \"\")", t("")),
        ("LEN(\"hello\")", n(5.0)),
    ]);
}

#[test]
fn datetime_functions() {
    check(&[
        ("DATE(2020, 1, 1)", n(43831.0)),
        ("DATE(2020, 13, 1)", n(44197.0)),
        ("DATE(1900, 1, 1)", n(1.0)),
        ("DATE(1900, 2, 29)", n(60.0)), // the phantom leap day
        ("DATE(1900, 3, 1)", n(61.0)),
        ("YEAR(43831)", n(2020.0)),
        ("MONTH(43831)", n(1.0)),
        ("DAY(43831)", n(1.0)),
        ("DAY(60)", n(29.0)),
        ("MONTH(60)", n(2.0)),
        ("DAY(59)", n(28.0)),
        ("WEEKDAY(43831)", n(4.0)),
        ("WEEKDAY(43831, 2)", n(3.0)),
        ("EDATE(43831, 1)", n(43862.0)),
        ("EOMONTH(43831, 0)", n(43861.0)),
        ("DATEDIF(43831, 44196, \"D\")", n(365.0)),
        ("DATEDIF(43831, 44196, \"M\")", n(11.0)),
        ("DATEDIF(43831, 44196, \"Y\")", n(0.0)),
        ("HOUR(0.5)", n(12.0)),
        ("HOUR(0.75)", n(18.0)),
        ("MINUTE(0.5)", n(0.0)),
        ("TIME(12, 0, 0)", n(0.5)),
        ("TODAY()", e(ErrorValue::Value)),
        ("NOW()", e(ErrorValue::Value)),
    ]);
    assert_eq!(eval_now("TODAY()"), n(43831.0));
    assert_eq!(eval_now("NOW()"), n(43831.5));
    approx("TIME(6, 0, 0)", 0.25);
}

#[test]
fn logical_functions() {
    check(&[
        ("IFERROR(1/0, \"x\")", t("x")),
        ("IFERROR(5, \"x\")", n(5.0)),
        ("IFNA(NA(), \"y\")", t("y")),
        ("IFNA(1/0, \"y\")", e(ErrorValue::Div0)),
        ("IFS(FALSE, 1, TRUE, 2)", n(2.0)),
        ("IFS(FALSE, 1, FALSE, 2)", e(ErrorValue::NA)),
        ("SWITCH(2, 1, \"a\", 2, \"b\", \"def\")", t("b")),
        ("SWITCH(9, 1, \"a\", \"def\")", t("def")),
        ("SWITCH(9, 1, \"a\")", e(ErrorValue::NA)),
        ("XOR(TRUE, FALSE)", b(true)),
        ("XOR(TRUE, TRUE)", b(false)),
        ("IF(TRUE, 1, 1/0)", n(1.0)),
        ("IFERROR(1, 1/0)", n(1.0)),
    ]);
}

#[test]
fn lookup_functions() {
    check(&[
        ("VLOOKUP(2, E1:F4, 2, FALSE)", t("two")),
        ("VLOOKUP(2.5, E1:F4, 2)", t("two")),
        ("VLOOKUP(9, E1:F4, 2, FALSE)", e(ErrorValue::NA)),
        ("HLOOKUP(3, H1:K2, 2, FALSE)", t("c")),
        ("INDEX(A1:A5, 3)", n(30.0)),
        ("INDEX(E1:F4, 2, 2)", t("two")),
        ("MATCH(30, A1:A5, 0)", n(3.0)),
        ("MATCH(35, A1:A5, 1)", n(3.0)),
        ("XLOOKUP(2, E1:E4, F1:F4)", t("two")),
        ("XLOOKUP(9, E1:E4, F1:F4, \"none\")", t("none")),
        ("CHOOSE(2, \"a\", \"b\", \"c\")", t("b")),
        ("ROW(A5)", n(5.0)),
        ("COLUMN(C1)", n(3.0)),
        ("ROWS(A1:A5)", n(5.0)),
        ("COLUMNS(E1:F4)", n(2.0)),
    ]);
}

/// without a calling cell the referenceless forms stay #VALUE!; `ROWS`/`COLUMNS`
/// have no referenceless form at all.
#[test]
fn referenceless_position_needs_a_calling_cell() {
    check(&[
        ("ROW()", e(ErrorValue::Value)),
        ("COLUMN()", e(ErrorValue::Value)),
        ("ROWS()", e(ErrorValue::Value)),
        ("COLUMNS()", e(ErrorValue::Value)),
        ("ROW(A1,B1)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn offset_shifts_a_single_cell() {
    check(&[
        ("OFFSET(A1, 2, 0)", n(30.0)),
        ("OFFSET(A1, 0, 2)", n(1.0)),
        ("OFFSET(A1, 1, 1)", t("banana")),
        ("OFFSET(C3, -2, -2)", n(10.0)),
        ("OFFSET($A$1, 4, 0)", n(50.0)),
        ("offset(A1, 2, 0)", n(30.0)),
    ]);
}

#[test]
fn offset_sizes_default_to_the_reference() {
    check(&[
        ("ROW(OFFSET(E1:F4, 1, 0))", n(2.0)),
        ("COLUMN(OFFSET(E1:F4, 0, 2))", n(7.0)),
        ("ROWS(OFFSET(E1:F4, 1, 0))", n(4.0)),
        ("COLUMNS(OFFSET(E1:F4, 1, 0))", n(2.0)),
        ("ROWS(OFFSET(A1, 1, 0))", n(1.0)),
        ("COLUMNS(OFFSET(A:B, 0, 1))", n(2.0)),
        ("COLUMN(OFFSET(A:B, 0, 1))", n(2.0)),
    ]);
}

#[test]
fn offset_resizes_and_extends_backwards() {
    check(&[
        ("SUM(OFFSET(A1, 1, 0, 3, 1))", n(90.0)),
        ("SUM(OFFSET(A1, 1, 0, 3))", n(90.0)),
        ("SUM(OFFSET(A3, 0, 0, -3, 1))", n(60.0)),
        ("ROW(OFFSET(A5, 0, 0, -3, 1))", n(3.0)),
        ("ROWS(OFFSET(A5, 0, 0, -3, 1))", n(3.0)),
        ("COLUMN(OFFSET(C1, 0, 0, 1, -3))", n(1.0)),
        ("COLUMNS(OFFSET(C1, 0, 0, 1, -3))", n(3.0)),
        ("ROWS(OFFSET(A1, 0, 0, 2, 3))", n(2.0)),
        ("COLUMNS(OFFSET(A1, 0, 0, 2, 3))", n(3.0)),
    ]);
}

#[test]
fn offset_rejects_empty_and_off_sheet_rectangles() {
    check(&[
        ("OFFSET(A1, 0, 0, 0, 1)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 0, 0, 1, 0)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 0, 0, Z9, 1)", e(ErrorValue::Ref)),
        ("OFFSET(A1, -1, 0)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 0, -1)", e(ErrorValue::Ref)),
        ("OFFSET(A2, 0, 0, -3, 1)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 1048576, 0)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 0, 16384)", e(ErrorValue::Ref)),
        ("OFFSET(A1, 0, 0, 1048577, 1)", e(ErrorValue::Ref)),
        ("SUM(OFFSET(A1, -1, 0))", e(ErrorValue::Ref)),
    ]);
}

#[test]
fn offset_argument_errors_and_scalar_context() {
    check(&[
        ("OFFSET(A1, 0, 0, 2, 1)", e(ErrorValue::Value)),
        ("OFFSET(5, 1, 1)", e(ErrorValue::Value)),
        ("OFFSET(A1, 1)", e(ErrorValue::Value)),
        ("OFFSET(A1, 1, 1, 1, 1, 1)", e(ErrorValue::Value)),
        ("OFFSET(A1, 1/0, 0)", e(ErrorValue::Div0)),
        ("OFFSET(A1, 0, 0, NA(), 1)", e(ErrorValue::NA)),
    ]);
}

#[test]
fn offset_feeds_the_other_reference_functions() {
    check(&[
        ("VLOOKUP(2, OFFSET(E1, 0, 0, 4, 2), 2, FALSE)", t("two")),
        ("MATCH(30, OFFSET(A1, 0, 0, 5, 1), 0)", n(3.0)),
        ("INDEX(OFFSET(A1, 0, 0, 5, 1), 4)", n(40.0)),
        ("SUM(OFFSET(OFFSET(A1, 1, 0), 1, 0, 2, 1))", n(70.0)),
        ("AVERAGE(OFFSET(A1, 0, 0, 5, 1))", n(30.0)),
    ]);
}

#[test]
fn info_functions() {
    check(&[
        ("ISBLANK(A6)", b(true)),
        ("ISBLANK(A1)", b(false)),
        ("ISNUMBER(A1)", b(true)),
        ("ISTEXT(B1)", b(true)),
        ("ISLOGICAL(TRUE)", b(true)),
        ("ISERROR(1/0)", b(true)),
        ("ISERR(1/0)", b(true)),
        ("ISERR(NA())", b(false)),
        ("ISNA(NA())", b(true)),
        ("NA()", e(ErrorValue::NA)),
        ("N(5)", n(5.0)),
        ("N(\"x\")", n(0.0)),
        ("N(TRUE)", n(1.0)),
    ]);
}

#[test]
fn case_insensitive_names() {
    check(&[
        ("sum(A1:A5)", n(150.0)),
        ("Vlookup(2, E1:F4, 2, false)", t("two")),
        ("mode.sngl(1, 2, 2)", n(2.0)),
        ("stdev.p(2, 4, 4, 4, 5, 5, 7, 9)", n(2.0)),
    ]);
}

#[test]
fn transpose_is_identity_on_single_values() {
    check(&[
        ("TRANSPOSE(A1)", n(10.0)),
        ("TRANSPOSE(B1)", t("apple")),
        ("TRANSPOSE(5)", n(5.0)),
        ("TRANSPOSE(\"hi\")", t("hi")),
        ("TRANSPOSE(TRUE)", b(true)),
        ("TRANSPOSE(2+3)", n(5.0)),
        ("TRANSPOSE(TRANSPOSE(A1))", n(10.0)),
        ("TRANSPOSE(A6)", n(0.0)), // blanks transpose to 0, not empty
        ("TRANSPOSE(1/0)", e(ErrorValue::Div0)),
        ("TRANSPOSE()", e(ErrorValue::Value)),
        ("TRANSPOSE(A1, A2)", e(ErrorValue::Value)),
    ]);
}

/// the engine has no array value type, so a multi-cell TRANSPOSE is #VALUE!
/// wherever it appears rather than flowing into the caller.
#[test]
fn transpose_of_a_multi_cell_area_has_no_representable_result() {
    check(&[
        ("TRANSPOSE(A1:A5)", e(ErrorValue::Value)),
        ("TRANSPOSE(H1:K1)", e(ErrorValue::Value)),
        ("TRANSPOSE(E1:F4)", e(ErrorValue::Value)),
        ("TRANSPOSE(A:A)", e(ErrorValue::Value)),
        ("SUM(TRANSPOSE(A1:A5))", e(ErrorValue::Value)),
        ("SUMPRODUCT(C1:C5, TRANSPOSE(A1:A5))", e(ErrorValue::Value)),
        ("ROWS(TRANSPOSE(E1:F4))", e(ErrorValue::Value)),
    ]);
}

/// matrix fixture: a 2x3 at M1:O2, a 3x2 at M4:N6, a 2x1 at M8:M9, and
/// one-row operands at M11:N11 (trailing text) and M13:N13 (trailing blank).
fn matrix_fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    let put = |s: &mut Sheet, a1: &str, v: CellValue| {
        s.set_cell(
            CellRef::parse_a1(a1).unwrap(),
            Cell {
                value: v,
                ..Cell::default()
            },
        );
    };
    for (a1, v) in [
        ("M1", 1.0),
        ("N1", 2.0),
        ("O1", 3.0),
        ("M2", 4.0),
        ("N2", 5.0),
        ("O2", 6.0),
        ("M4", 7.0),
        ("N4", 8.0),
        ("M5", 9.0),
        ("N5", 10.0),
        ("M6", 11.0),
        ("N6", 12.0),
        ("M8", 1.0),
        ("M9", 1.0),
        ("M11", 1.0),
        ("M13", 1.0),
    ] {
        put(&mut s, a1, n(v));
    }
    put(&mut s, "N11", t("x"));
    put(&mut s, "N15", b(true));
    put(&mut s, "M15", n(1.0));
    wb.sheets.push(s);
    wb
}

fn eval_matrix(src: &str) -> CellValue {
    let wb = matrix_fixture();
    let expr = parse_formula(src).expect("parse");
    let ctx = EvalContext::new(&wb, SheetId(0));
    evaluate(&expr, &ctx)
}

/// M1:O2 times M4:N6 is [[58, 64], [139, 154]]; each element is checked by
/// multiplying the matching row and column, so every product is covered.
#[test]
fn mmult_multiplies_a_2x3_by_a_3x2() {
    for (src, expected) in [
        ("MMULT(M1:O1, M4:M6)", 58.0),
        ("MMULT(M1:O1, N4:N6)", 64.0),
        ("MMULT(M2:O2, M4:M6)", 139.0),
        ("MMULT(M2:O2, N4:N6)", 154.0),
    ] {
        assert_eq!(eval_matrix(src), n(expected), "{src}");
    }
}

/// the engine stores one value per cell, so a wider product yields its
/// top-left element: what excel caches in the array formula's anchor cell.
#[test]
fn mmult_returns_the_top_left_element_of_a_wider_product() {
    assert_eq!(eval_matrix("MMULT(M1:O2, M4:N6)"), n(58.0));
}

#[test]
fn mmult_rejects_mismatched_and_non_numeric_operands() {
    for src in [
        "MMULT(M1:O2, M1:O2)",
        "MMULT(M11:N11, M8:M9)",
        "MMULT(M13:N13, M8:M9)",
        "MMULT(M15:N15, M8:M9)",
        "MMULT(M1:O1)",
        "MMULT(M1:O1, M4:M6, M4:M6)",
    ] {
        assert_eq!(eval_matrix(src), e(ErrorValue::Value), "{src}");
    }
}

/// an argument that is not a reference is a 1x1 matrix, and an argument that
/// evaluates to an error propagates it -- so an unsupported inner function
/// still surfaces `#NAME?` rather than being masked as `#VALUE!`.
#[test]
fn mmult_handles_scalar_and_erroring_arguments() {
    assert_eq!(eval_matrix("MMULT(3, 4)"), n(12.0));
    assert_eq!(eval_matrix("MMULT(1/0, M8:M9)"), e(ErrorValue::Div0));
    assert_eq!(eval_matrix("MMULT(M1:O1, NOSUCH())"), e(ErrorValue::Name));
    assert_eq!(eval_matrix("MMULT(NOSUCH(), M8:M9)"), e(ErrorValue::Name));
}

fn draws(src: &str, seed: Option<u64>, count: usize) -> Vec<f64> {
    let wb = fixture();
    let expr = parse_formula(src).expect("parse");
    let mut ctx = EvalContext::new(&wb, SheetId(0));
    ctx.rand_seed = seed;
    (0..count)
        .map(|_| match evaluate(&expr, &ctx) {
            CellValue::Number { value } => value,
            other => panic!("formula {src:?}: expected number, got {other:?}"),
        })
        .collect()
}

/// the rounding and error rules here were measured against Excel for Mac over
/// 400 draws per case: `bottom > top` errors before any rounding, the draw
/// spans `ceil(bottom)..=floor(top)`, and an empty span yields `ceil(bottom)`.
#[test]
fn randbetween_matches_excels_rounding() {
    check(&[
        ("RANDBETWEEN(5, 5)", n(5.0)),
        ("RANDBETWEEN(1.8, 2.2)", n(2.0)),
        ("RANDBETWEEN(2.9, 3.1)", n(3.0)),
        ("RANDBETWEEN(-0.5, 0.5)", n(0.0)),
        ("RANDBETWEEN(1.5, 1.6)", n(2.0)),
        ("RANDBETWEEN(2.2, 2.2)", n(3.0)),
        ("RANDBETWEEN(-1.5, -1.4)", n(-1.0)),
        ("RANDBETWEEN(0.1, 0.9)", n(1.0)),
        ("RANDBETWEEN(-0.9, -0.1)", n(0.0)),
        ("RANDBETWEEN(2.5, 2.1)", e(ErrorValue::Num)),
        ("RANDBETWEEN(3, 1)", e(ErrorValue::Num)),
        ("RANDBETWEEN(1)", e(ErrorValue::Value)),
        ("RANDBETWEEN(1, 2, 3)", e(ErrorValue::Value)),
        ("RANDBETWEEN(\"x\", 2)", e(ErrorValue::Value)),
        ("RANDBETWEEN(A1, A1)", n(10.0)),
    ]);
}

#[test]
fn randbetween_covers_its_range_and_never_leaves_it() {
    for (src, want) in [
        ("RANDBETWEEN(1.2, 3.8)", vec![2.0, 3.0]),
        ("RANDBETWEEN(-3.5, -1.2)", vec![-3.0, -2.0]),
        ("RANDBETWEEN(-1, 1)", vec![-1.0, 0.0, 1.0]),
    ] {
        let mut seen: Vec<f64> = draws(src, None, 2_000);
        for value in &seen {
            assert!(want.contains(value), "formula {src:?} drew {value}");
        }
        seen.sort_by(f64::total_cmp);
        seen.dedup();
        assert_eq!(seen, want, "formula {src:?} never covered its range");
    }
}

#[test]
fn randbetween_replays_a_pinned_seed() {
    let src = "RANDBETWEEN(1, 1000000)";
    assert_eq!(draws(src, Some(7), 16), draws(src, Some(7), 16));
    assert_ne!(draws(src, Some(7), 16), draws(src, Some(8), 16));
    assert_ne!(draws(src, None, 16), draws(src, None, 16));
}
/// an argument that cannot become an area may still have said why: OFFSET
/// past the sheet edge is #REF!, and the count must not flatten it to #VALUE!.
#[test]
fn reference_counts_propagate_their_arguments_error() {
    check(&[
        ("ROWS(OFFSET(A1,-1,0))", e(ErrorValue::Ref)),
        ("COLUMNS(OFFSET(A1,0,-1))", e(ErrorValue::Ref)),
        ("ROWS(1/0)", e(ErrorValue::Div0)),
        ("ROWS(5)", e(ErrorValue::Value)),
    ]);
}

/// excel stores post-2007 functions with an `_xlfn.` prefix (`_xlfn._xlws.`
/// for worksheet-only ones), so the prefix must resolve to the same builtin.
/// an array builtin called outside an array formula shows its top-left value;
/// a prefixed name we do not implement stays `#NAME?`.
#[test]
fn xlfn_prefixed_names_resolve_to_the_same_builtin() {
    check(&[
        ("_xlfn.TEXTJOIN(\"-\", TRUE, \"a\", \"\", \"b\")", t("a-b")),
        (
            "_XLFN.TEXTJOIN(\"-\", FALSE, \"a\", \"\", \"b\")",
            t("a--b"),
        ),
        ("_xlfn.CONCAT(\"a\", \"b\")", t("ab")),
        ("_xlfn.IFNA(1/0, 7)", e(ErrorValue::Div0)),
        ("_xlfn.IFNA(NA(), 7)", n(7.0)),
        ("_xlfn.XLOOKUP(2, E1:E4, F1:F4)", t("two")),
        ("_xlfn._xlws.FILTER(A1:A5, C1:C5)", n(10.0)),
        ("_xlfn.LET(_xlpm.x, 6, _xlpm.x * 7)", n(42.0)),
        ("_xlfn.NOSUCH()", e(ErrorValue::Name)),
        ("_xlws.CONCAT(\"a\", \"b\")", e(ErrorValue::Name)),
    ]);
}
