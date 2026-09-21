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

/// A name the workbook does not define is #NAME? wherever it stands, so a
/// function that wanted a reference there reports that rather than #VALUE!.
#[test]
fn an_unknown_name_in_a_reference_argument_is_a_name_error() {
    check(&[
        ("nosuchname", e(ErrorValue::Name)),
        ("VLOOKUP(2,nosuchname,2,FALSE)", e(ErrorValue::Name)),
        ("HLOOKUP(2,nosuchname,2,FALSE)", e(ErrorValue::Name)),
        ("MATCH(2,nosuchname,0)", e(ErrorValue::Name)),
        ("XMATCH(2,nosuchname)", e(ErrorValue::Name)),
        ("XLOOKUP(2,nosuchname,E1:E4)", e(ErrorValue::Name)),
        ("XLOOKUP(2,E1:E4,nosuchname)", e(ErrorValue::Name)),
        ("INDEX(nosuchname,1,1)", e(ErrorValue::Name)),
        ("OFFSET(nosuchname,0,0)", e(ErrorValue::Name)),
        ("SUMIF(nosuchname,\">1\")", e(ErrorValue::Name)),
        ("VLOOKUP(2,E1:F4,9,FALSE)", e(ErrorValue::Ref)),
        ("VLOOKUP(2,\"x\",2,FALSE)", e(ErrorValue::Value)),
    ]);
}

/// Excel's lotus-style leading `+` passes its operand through, so `+A1` on a
/// text cell is that text and `+J1&+K1` joins them; only negation wants a
/// number.
#[test]
fn a_leading_plus_passes_its_operand_through() {
    check(&[
        ("+B1", t("apple")),
        ("+B1&+B2", t("applebanana")),
        ("+A1", n(10.0)),
        ("+A1+1", n(11.0)),
        ("-B1", e(ErrorValue::Value)),
        ("-A1", n(-10.0)),
    ]);
}

/// Text that spells a date reads as its serial wherever a number is wanted,
/// not only inside the date functions.
#[test]
fn date_text_coerces_wherever_a_number_is_wanted() {
    check(&[
        ("CONVERT(\"1/8/2020\",\"in\",\"m\")", n(43_838.0 * 0.0254)),
        ("ABS(\"1/8/2020\")", n(43_838.0)),
        ("CONVERT(\"not a date\",\"in\",\"m\")", e(ErrorValue::Value)),
    ]);
    assert_eq!(eval_now("MONTH(\"Nov\"&1)"), n(11.0));
    assert_eq!(eval_now("DAY(\"Nov\"&1)"), n(1.0));
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

#[test]
fn match_accepts_a_computed_block() {
    check(&[
        ("MATCH(\"3three\",E1:E4&F1:F4,0)", n(3.0)),
        ("INDEX(F1:F4,MATCH(\"3three\",E1:E4&F1:F4,0))", t("three")),
        ("MATCH(\"nope\",E1:E4&F1:F4,0)", e(ErrorValue::NA)),
        ("MATCH(6,C1:C5*2,0)", n(3.0)),
        ("MATCH(2,{1,2,3},0)", n(2.0)),
    ]);
}

#[test]
fn the_range_operator_spans_computed_endpoints() {
    check(&[
        ("SUM(A1:INDEX(A1:A5,3))", n(60.0)),
        ("SUM(INDEX(A1:A5,2):INDEX(A1:A5,4))", n(90.0)),
        ("COUNT(A1:OFFSET(A1,2,0))", n(3.0)),
        ("MIN(A1:INDEX(A1:A5,MATCH(99,A1:A5,0)))", e(ErrorValue::NA)),
    ]);
    assert!(parse_formula("A1:\"x\"").is_err());
    assert!(parse_formula("SUM(A1:INDEX(A1:A5,3))").is_ok());
}

#[test]
fn sumproduct_reads_computed_arguments_as_blocks() {
    check(&[
        ("SUMPRODUCT(--(B1:B5=\"apple\"))", n(3.0)),
        ("SUMPRODUCT(--(B1:B5=\"apple\"),A1:A5)", n(90.0)),
        ("SUMPRODUCT((B1:B5=\"apple\")*A1:A5)", n(90.0)),
        ("SUMPRODUCT(A1:A5,C1:C5)", n(550.0)),
        (
            "SUMPRODUCT(--(B1:B5=\"apple\"),A1:A4)",
            e(ErrorValue::Value),
        ),
    ]);
}

#[test]
fn time_text_coerces_to_a_fraction_of_a_day() {
    approx("MROUND(0.37612268519,\"0:15\")", 0.375);
    approx("\"1:30\"*24", 1.5);
    approx("\"0:15\"+0", 0.010_416_666_666_666_666);
    approx("\"9:00 PM\"+0", 0.875);
    approx("\"12:00 AM\"+0", 0.0);
    approx("\"36:00\"+0", 1.5);
    approx("\"1:02:03\"+0", 0.043_090_277_777_777_78);
    check(&[
        ("\"1:60\"+0", e(ErrorValue::Value)),
        ("\"13:00 PM\"+0", e(ErrorValue::Value)),
        ("\"x:00\"+0", e(ErrorValue::Value)),
    ]);
}

#[test]
fn boolean_literals_also_spell_as_calls() {
    check(&[
        ("TRUE()", b(true)),
        ("FALSE()", b(false)),
        ("TRUE", b(true)),
        ("NOT(FALSE())", b(true)),
        ("MATCH(TRUE(),{FALSE,TRUE},0)", n(2.0)),
        ("TRUE(1)", e(ErrorValue::Value)),
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
        // excel calls serial 0 "january 0, 1900", which a blank date reads as
        ("YEAR(0)", n(1900.0)),
        ("MONTH(0)", n(1.0)),
        ("DAY(0)", n(0.0)),
        ("DATEDIF(0,41115,\"y\")", n(112.0)),
        ("DATEDIF(0,0,\"y\")", n(0.0)),
        // NPV discounts from the end of period one
        ("NPV(0.1,100)", n(100.0 / 1.1)),
        ("NPV(0,10,20,30)", n(60.0)),
        ("NPV(-1,10)", e(ErrorValue::Num)),
        ("NPV(0.1)", e(ErrorValue::Value)),
        // a date function reads text where a date is wanted
        ("MONTH(\"2024-03-05\")", n(3.0)),
        ("YEAR(\"2024-03-05\")", n(2024.0)),
        ("DAY(\"March 5, 2024\")", n(5.0)),
        ("MONTH(\"not a date\")", e(ErrorValue::Value)),
        // the plain workday pair is the .INTL one with a fixed weekend
        ("NETWORKDAYS(45292,45303)", n(10.0)),
        ("NETWORKDAYS(45303,45292)", n(-10.0)),
        ("WORKDAY(45292,5)", n(45299.0)),
        ("WORKDAY(45292,-1)", n(45289.0)),
        ("NETWORKDAYS(45292)", e(ErrorValue::Value)),
        // week 1 holds jan 1, so the count depends on which weekday that is:
        // 2005-01-01 is a saturday, 2023-01-01 a sunday, 2024-01-01 a monday
        ("WEEKNUM(38580)", n(34.0)),
        ("WEEKNUM(38580, 2)", n(34.0)),
        ("WEEKNUM(45291)", n(53.0)),
        ("WEEKNUM(45292)", n(1.0)),
        ("WEEKNUM(45292, 2)", n(1.0)),
        ("WEEKNUM(46023, 99)", e(ErrorValue::Num)),
        ("WEEKNUM(44830, 1)", n(40.0)),
        ("WEEKNUM(44830, 2)", n(40.0)),
        ("WEEKNUM(44562, 1)", n(1.0)),
        ("WEEKNUM(44563, 1)", n(2.0)),
        ("WEEKNUM(44197, 1)", n(1.0)),
        ("WEEKNUM(44199, 1)", n(2.0)),
        ("WEEKNUM(0, 1)", n(0.0)),
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
        ("TEXTBEFORE(\"a-b\", \"X\", 1, 0, 0, \"none\")", t("none")),
        ("TEXTBEFORE(\"a-b\", \"X\", 1, 0, 1)", t("a-b")),
        ("TEXTAFTER(\"a-b\", \"X\", 1, 0, 1)", t("")),
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

/// a multi-cell TRANSPOSE has no single-cell value, so reading one as a scalar
/// is #VALUE!; an aggregate consumes the block instead.
#[test]
fn transpose_of_a_multi_cell_area_aggregates_but_has_no_scalar_value() {
    check(&[
        ("TRANSPOSE(A1:A5)", e(ErrorValue::Value)),
        ("TRANSPOSE(H1:K1)", e(ErrorValue::Value)),
        ("TRANSPOSE(E1:F4)", e(ErrorValue::Value)),
        ("TRANSPOSE(A:A)", e(ErrorValue::Value)),
        ("SUM(TRANSPOSE(A1:A5))", n(150.0)),
        ("MEDIAN(TRANSPOSE(A1:A5))", n(30.0)),
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

/// DATEVALUE reads the date orders a US locale writes, ignores a trailing
/// time, and refuses anything that is not a date.
#[test]
fn datevalue_reads_written_dates() {
    check(&[
        ("DATEVALUE(\"3/28/2001\")", n(36978.0)),
        ("DATEVALUE(\"03/28/01\")", n(36978.0)),
        ("DATEVALUE(\"2001-03-28\")", n(36978.0)),
        ("DATEVALUE(\"28-Mar-2001\")", n(36978.0)),
        ("DATEVALUE(\"March 28, 2001\")", n(36978.0)),
        ("DATEVALUE(\"Mar 28, 01\")", n(36978.0)),
        ("DATEVALUE(\"1/1/2020 13:30\")", n(43831.0)),
        ("DATEVALUE(\"1/1/1900\")", n(1.0)),
        ("DATEVALUE(\"2/29/1900\")", n(60.0)),
        ("DATEVALUE(TEXT(36978, \"mm/dd/yy\"))", n(36978.0)),
        ("DATEVALUE(\"1/1/35\")", n(12785.0)),
        ("DATEVALUE(\"hello\")", e(ErrorValue::Value)),
        ("DATEVALUE(\"13/1/2001\")", e(ErrorValue::Value)),
        ("DATEVALUE(\"2/30/2001\")", e(ErrorValue::Value)),
        ("DATEVALUE(\"12:00\")", e(ErrorValue::Value)),
        ("DATEVALUE(\"1/1/1899\")", e(ErrorValue::Value)),
        ("DATEVALUE(43831)", e(ErrorValue::Value)),
        ("DATEVALUE(\"1/1/2020\", 1)", e(ErrorValue::Value)),
        ("DATEVALUE(1/0)", e(ErrorValue::Div0)),
    ]);
    // a month and a day alone need a clock to know which year they belong to
    assert_eq!(eval("DATEVALUE(\"3/28\")"), e(ErrorValue::Value));
    assert_eq!(eval_now("DATEVALUE(\"1/1\")"), n(43831.0));
}

/// YEARFRAC's five day-count bases, including the 30/360 rules that only
/// differ on month ends.
#[test]
fn yearfrac_counts_days_on_every_basis() {
    check(&[
        ("YEARFRAC(DATE(2020,1,1), DATE(2021,1,1))", n(1.0)),
        ("YEARFRAC(DATE(2021,1,1), DATE(2020,1,1))", n(1.0)),
        ("YEARFRAC(DATE(2020,1,1), DATE(2021,1,1), 1)", n(1.0)),
        ("YEARFRAC(DATE(2020,1,1), DATE(2020,12,31), 3)", n(1.0)),
        (
            "YEARFRAC(DATE(2020,1,1), DATE(2020,2,1), 5)",
            e(ErrorValue::Num),
        ),
        ("YEARFRAC(DATE(2020,1,1))", e(ErrorValue::Value)),
        ("YEARFRAC(1/0, 2)", e(ErrorValue::Div0)),
    ]);
    approx(
        "YEARFRAC(DATE(2020,1,31), DATE(2020,3,31), 0)",
        60.0 / 360.0,
    );
    approx(
        "YEARFRAC(DATE(2020,2,29), DATE(2020,3,31), 0)",
        30.0 / 360.0,
    );
    approx(
        "YEARFRAC(DATE(2020,2,29), DATE(2020,3,31), 4)",
        31.0 / 360.0,
    );
    approx(
        "YEARFRAC(DATE(2020,1,1), DATE(2020,12,31), 2)",
        365.0 / 360.0,
    );
    approx("YEARFRAC(DATE(2019,1,1), DATE(2019,7,1), 1)", 181.0 / 365.0);
}

/// the `.INTL` workday pair, over both weekend encodings. 2020-01-01 is a
/// wednesday, so 2020-01-04 and 2020-01-05 are the weekend of that week.
#[test]
fn intl_workday_functions_read_both_weekend_encodings() {
    check(&[
        ("NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7))", n(5.0)),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), \"0000011\")",
            n(5.0),
        ),
        ("NETWORKDAYS.INTL(DATE(2020,1,7), DATE(2020,1,1))", n(-5.0)),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), 11)",
            n(6.0),
        ),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), 1, DATE(2020,1,2))",
            n(4.0),
        ),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), 1, DATE(2020,1,4))",
            n(5.0),
        ),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), \"1111111\")",
            n(0.0),
        ),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), \"012\")",
            e(ErrorValue::Value),
        ),
        (
            "NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), 8)",
            e(ErrorValue::Num),
        ),
        ("NETWORKDAYS.INTL(DATE(2020,1,1))", e(ErrorValue::Value)),
        ("WORKDAY.INTL(DATE(2020,1,3), 1)", n(43836.0)),
        ("WORKDAY.INTL(DATE(2020,1,3), 1, \"0000000\")", n(43834.0)),
        ("WORKDAY.INTL(DATE(2020,1,6), -1)", n(43833.0)),
        ("WORKDAY.INTL(DATE(2020,1,3), 0)", n(43833.0)),
        ("WORKDAY.INTL(DATE(2020,1,3), 10)", n(43847.0)),
        (
            "WORKDAY.INTL(DATE(2020,1,3), 1, 1, DATE(2020,1,6))",
            n(43837.0),
        ),
        (
            "WORKDAY.INTL(DATE(2020,1,3), 1, \"1111111\")",
            e(ErrorValue::Value),
        ),
        ("WORKDAY.INTL(DATE(2020,1,3))", e(ErrorValue::Value)),
    ]);
    // the weekend argument written as a gap falls back to saturday and sunday
    assert_eq!(
        eval("NETWORKDAYS.INTL(DATE(2020,1,1), DATE(2020,1,7), \"0000011\",)"),
        n(5.0)
    );
}

/// ADDRESS in both notations, with every absolute/relative combination.
#[test]
fn address_writes_a_reference_as_text() {
    check(&[
        ("ADDRESS(1, 1)", t("$A$1")),
        ("ADDRESS(2, 3)", t("$C$2")),
        ("ADDRESS(2, 3, 2)", t("C$2")),
        ("ADDRESS(2, 3, 3)", t("$C2")),
        ("ADDRESS(2, 3, 4)", t("C2")),
        ("ADDRESS(35, 1)", t("$A$35")),
        ("ADDRESS(1, 27)", t("$AA$1")),
        ("ADDRESS(2, 3, 1, FALSE)", t("R2C3")),
        ("ADDRESS(2, 3, 4, FALSE)", t("R[2]C[3]")),
        ("ADDRESS(1, 1, 1, TRUE, \"Sheet1\")", t("Sheet1!$A$1")),
        ("ADDRESS(1, 1, 1, TRUE, \"My Sheet\")", t("'My Sheet'!$A$1")),
        ("ADDRESS(0, 1)", e(ErrorValue::Value)),
        ("ADDRESS(1, 0)", e(ErrorValue::Value)),
        ("ADDRESS(1, 16385)", e(ErrorValue::Value)),
        ("ADDRESS(1, 1, 5)", e(ErrorValue::Value)),
        ("ADDRESS(1)", e(ErrorValue::Value)),
        ("ADDRESS(1/0, 1)", e(ErrorValue::Div0)),
        ("INDIRECT(ADDRESS(2, 1))", n(20.0)),
    ]);
}

/// HYPERLINK shows the caption, but an error in the jump target still wins.
#[test]
fn hyperlink_shows_its_caption() {
    check(&[
        ("HYPERLINK(\"http://x\", \"go\")", t("go")),
        ("HYPERLINK(\"http://x\")", t("http://x")),
        ("HYPERLINK(\"#a\", 5)", n(5.0)),
        ("HYPERLINK(1/0, \"go\")", e(ErrorValue::Div0)),
        ("HYPERLINK(\"#a\", 1/0)", e(ErrorValue::Div0)),
        ("HYPERLINK(\"a\", \"b\", \"c\")", e(ErrorValue::Value)),
    ]);
}

/// FV over both payment timings, the zero-rate shortcut, and its error cases.
#[test]
fn fv_discounts_an_annuity() {
    check(&[
        ("FV(0, 10, -100)", n(1000.0)),
        ("FV(0, 10, -100, -500)", n(1500.0)),
        ("FV(0.1, 2)", e(ErrorValue::Value)),
        ("FV(0.1, 2, 0, -100, 1, 9)", e(ErrorValue::Value)),
        ("FV(1/0, 1, 1)", e(ErrorValue::Div0)),
    ]);
    approx("FV(0.05, 10, 0, -1000)", 1628.894626777442);
    approx("FV(0.05, 10, -100, 0)", 1_257.789_253_554_884);
    approx("FV(0.05, 10, -100, 0, 1)", 1320.6787162326282);
    approx("FV(0.05/12, 120, -100, -1000)", 17175.237442257);
}

/// NORM.DIST in both modes; the cumulative branch must hold to full double
/// precision because the corpus subtracts neighbouring values.
#[test]
fn norm_dist_covers_density_and_cumulative() {
    check(&[
        ("NORM.DIST(1, 0, 0, TRUE)", e(ErrorValue::Num)),
        ("NORM.DIST(1, 0, -1, TRUE)", e(ErrorValue::Num)),
        ("NORM.DIST(1, 0, 1)", e(ErrorValue::Value)),
        ("_xlfn.NORM.DIST(0, 0, 1, TRUE)", n(0.5)),
    ]);
    approx("NORM.DIST(0, 0, 1, FALSE)", 0.3989422804014327);
    approx("NORM.DIST(1.96, 0, 1, TRUE)", 0.9750021048517795);
    approx("NORM.DIST(-1.96, 0, 1, TRUE)", 0.024997895148220435);
    approx("NORM.DIST(-8, 0, 1, TRUE)", 6.220960574271786e-16);
    approx("NORM.DIST(42, 40, 1.5, TRUE)", 0.9087887802741321);
    approx("NORM.DIST(42, 40, 1.5, FALSE)", 0.10934004978399577);
}

/// the `-A` aggregates count text as zero and logicals as one or zero, where
/// the plain forms skip both. B1:B5 holds fruit names.
#[test]
fn a_suffixed_aggregates_count_text_and_logicals() {
    check(&[
        ("AVERAGE(B1:B5)", e(ErrorValue::Div0)),
        ("AVERAGEA(B1:B5)", n(0.0)),
        ("AVERAGEA(A1:A5)", n(30.0)),
        ("AVERAGEA(A1:B5)", n(15.0)),
        ("AVERAGEA(B1:B1)", n(0.0)),
        ("AVERAGEA()", e(ErrorValue::Div0)),
        ("AVERAGEA(\"x\")", e(ErrorValue::Value)),
        ("AVERAGEA(TRUE, TRUE, 4)", n(2.0)),
        ("MAXA(B1:B5)", n(0.0)),
        ("MAXA(A1:B5)", n(50.0)),
        ("MINA(A1:B5)", n(0.0)),
        ("MIN(A1:B5)", n(10.0)),
        ("MAXA(1/0)", e(ErrorValue::Div0)),
        ("STDEVA(A1:A5)", n(15.811388300841896)),
        ("STDEVA(5)", e(ErrorValue::Div0)),
    ]);
    approx("STDEVA(A1:B5)", 19.0029237516523);
}

/// the remaining aggregates, over C1:C5 = 1..5 and A1:A5 = 10..50.
#[test]
fn descriptive_aggregates() {
    check(&[
        ("SUMSQ(1, 2, 3)", n(14.0)),
        ("SUMSQ(C1:C5)", n(55.0)),
        ("SUMSQ()", n(0.0)),
        ("SUMSQ(1/0)", e(ErrorValue::Div0)),
        ("DEVSQ(C1:C5)", n(10.0)),
        ("DEVSQ(B1:B5)", e(ErrorValue::Num)),
        ("AVEDEV(C1:C5)", n(1.2)),
        ("AVEDEV(B1:B5)", e(ErrorValue::Num)),
        ("GEOMEAN(1, 4)", n(2.0)),
        ("GEOMEAN(1, 0)", e(ErrorValue::Num)),
        ("GEOMEAN(-1, 4)", e(ErrorValue::Num)),
        ("GEOMEAN(B1:B5)", e(ErrorValue::Num)),
        ("HARMEAN(1, 2, 4)", n(12.0 / 7.0)),
        ("HARMEAN(1, 0)", e(ErrorValue::Num)),
        ("SKEW(C1:C5)", n(0.0)),
        ("SKEW(1, 2)", e(ErrorValue::Div0)),
        ("SKEW(3, 3, 3)", e(ErrorValue::Div0)),
        ("KURT(1, 2, 3)", e(ErrorValue::Div0)),
        ("TRIMMEAN(C1:C5, 0.4)", n(3.0)),
        ("TRIMMEAN(C1:C5, 0)", n(3.0)),
        ("TRIMMEAN(C1:C5, 1)", e(ErrorValue::Num)),
        ("TRIMMEAN(C1:C5, -0.1)", e(ErrorValue::Num)),
        ("TRIMMEAN(B1:B5, 0.2)", e(ErrorValue::Num)),
        ("TRIMMEAN(C1:C5)", e(ErrorValue::Value)),
    ]);
    approx("GEOMEAN(A1:A5)", 26.051710846973528);
    approx("KURT(1, 2, 3, 10)", 3.228);
    approx("TRIMMEAN({1,2,3,4,100}, 0.4)", 3.0);
    approx("SUMSQ({1,2,3})", 14.0);
    approx("AVEDEV({1,2,3,4})", 1.0);
}

/// the range operator spans two references that only resolve at evaluation
/// time. A1:A5 = 10..50, C1:C5 = 1..5.
#[test]
fn range_operator_spans_resolved_endpoints() {
    check(&[
        ("SUM(A1:INDEX(A1:A5,3))", n(60.0)),
        ("MIN(A1:INDEX(A1:A5,3))", n(10.0)),
        ("SUM(INDEX(A1:A5,2):INDEX(A1:A5,4))", n(90.0)),
        ("SUM(A1:OFFSET(A1,2,0))", n(60.0)),
        ("SUM($A$1:OFFSET($A$1,,))", n(10.0)),
        ("COUNT(A1:INDEX(A:A,5))", n(5.0)),
        ("ROWS(A1:INDEX(A1:A5,3))", n(3.0)),
        ("SUM(A1:INDEX(A1:C5,3,2))", n(60.0)),
        ("INDEX(A1:A5,1):INDEX(A1:A5,1)", n(10.0)),
        // an end that cannot resolve reports its own error, not #VALUE!
        ("SUM(A1:INDEX(A1:A5,MATCH(99,C1:C5,0)))", e(ErrorValue::NA)),
        (
            "IFERROR(MIN(A1:INDEX(A1:A5,MATCH(99,C1:C5,0))),\"x\")",
            t("x"),
        ),
        ("SUM(A1:INDEX(A1:A5,99))", e(ErrorValue::Ref)),
        ("SUM(A1:#REF!)", e(ErrorValue::Ref)),
        ("SUM(A1:OFFSET(A1,-1,0))", e(ErrorValue::Ref)),
        ("SUM(A1:SUM(1))", e(ErrorValue::Ref)),
    ]);
}

/// a join whose ends sit on different sheets designates nothing.
#[test]
fn range_operator_refuses_two_sheets() {
    let mut workbook = Workbook::default();
    let mut first = Sheet::new("Sheet1");
    first.set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: n(7.0),
            ..Cell::default()
        },
    );
    workbook.sheets.push(first);
    workbook.sheets.push(Sheet::new("Other"));
    let context = EvalContext::new(&workbook, SheetId(0));
    let value = |src: &str| evaluate(&parse_formula(src).unwrap(), &context);
    assert_eq!(value("SUM(A1:Other!B2)"), e(ErrorValue::Ref));
    assert_eq!(
        value("SUM(Sheet1!A1:INDEX(Other!A1:A9,2))"),
        e(ErrorValue::Ref)
    );
    assert_eq!(value("SUM(Sheet1!A1:INDEX(Sheet1!A1:A9,2))"), n(7.0));
}

/// INDEX and OFFSET resolve references inside array formulas too, so an
/// index scalar evaluation cannot read is retried elementwise. B1:B5 holds
/// fruit names of length 5, 6, 5, 6, 5.
#[test]
fn reference_indexes_fall_back_to_array_evaluation() {
    check(&[
        ("LEN(B1:B5)", e(ErrorValue::Value)),
        ("INDEX(A1:A5,MATCH(6,LEN(B1:B5),0))", n(20.0)),
        ("OFFSET(A1,MATCH(6,LEN(B1:B5),0),0)", n(30.0)),
        ("SUM(A1:INDEX(A1:A5,MATCH(6,LEN(B1:B5),0)))", n(30.0)),
        // a reference index keeps its scalar answer rather than its first cell
        ("INDEX(A1:A5,C1:C3)", e(ErrorValue::Value)),
        ("INDEX(A1:A5,1/0)", e(ErrorValue::Div0)),
    ]);
}

/// the annuity family, over both payment timings and the zero-rate shortcut.
#[test]
fn annuity_functions() {
    check(&[
        ("PMT(0, 10, 1000)", n(-100.0)),
        ("PMT(0, 10, 1000, 500)", n(-150.0)),
        ("PMT(0.05, 0, 1000)", e(ErrorValue::Num)),
        ("PMT(0.05, 10)", e(ErrorValue::Value)),
        ("PV(0, 10, -100)", n(1000.0)),
        ("NPER(0, -100, 1000)", n(10.0)),
        ("NPER(0, 0, 1000)", e(ErrorValue::Num)),
        ("NPV(-1, 100)", e(ErrorValue::Num)),
        ("NPV(0.05)", e(ErrorValue::Value)),
        ("PMT(1/0, 1, 1)", e(ErrorValue::Div0)),
    ]);
    approx("PMT(0.04/12, 12, 5000)", -425.7495209777896);
    approx("PMT(0.05, 10, 1000, 0, 1)", -123.337690443292);
    approx("PV(0.05, 10, -100)", 772.1734929184817);
    approx("NPER(0.05, -100, 1000)", 14.206699082890461);
    approx("NPV(0.05, 100, 200, 300)", 535.795270489148);
    approx("NPV(0.05, {100,200,300})", 535.795270489148);
    // a loan pays itself off: the payment FV discounts back to the principal
    approx("FV(0.04/12, 12, PMT(0.04/12, 12, 5000), 5000)", 0.0);
}

/// TIMEVALUE reads a written time, drops the date in front of it, and keeps
/// only the fraction of a day.
#[test]
fn timevalue_reads_a_written_clock() {
    check(&[
        ("TIMEVALUE(\"2:24 AM\")", n(0.1)),
        ("TIMEVALUE(\"12:00 PM\")", n(0.5)),
        ("TIMEVALUE(\"12:00 AM\")", n(0.0)),
        ("TIMEVALUE(\"hello\")", e(ErrorValue::Value)),
        ("TIMEVALUE(\"1/1/2020\")", e(ErrorValue::Value)),
        ("TIMEVALUE(\"13:00 PM\")", e(ErrorValue::Value)),
        ("TIMEVALUE(0.5)", e(ErrorValue::Value)),
        ("TIMEVALUE(\"2:24 AM\", 1)", e(ErrorValue::Value)),
        ("TIMEVALUE(1/0)", e(ErrorValue::Div0)),
    ]);
    approx("TIMEVALUE(\"0:15\")", 0.010416666666666666);
    approx("TIMEVALUE(\"12:30:45\")", 0.5213541666666667);
    approx("TIMEVALUE(\"22-Aug-2011 6:35 AM\")", 0.2743055555555556);
    approx("TIMEVALUE(\"25:00\")", 0.041666666666666664);
    approx("TIMEVALUE(\"12:30\" & \" \" & \"PM\")", 0.5208333333333334);
}

/// DAYS counts whole days either way round, and reads a written date.
#[test]
fn days_between_two_dates() {
    check(&[
        ("DAYS(DATE(2020,3,1), DATE(2020,2,1))", n(29.0)),
        ("DAYS(DATE(2020,2,1), DATE(2020,3,1))", n(-29.0)),
        ("DAYS(43831, 43831)", n(0.0)),
        ("DAYS(\"3/1/2021\", \"2/1/2021\")", n(28.0)),
        ("DAYS(43831.9, 43830.1)", n(1.0)),
        ("DAYS(\"hello\", 1)", e(ErrorValue::Value)),
        ("DAYS(-1, 1)", e(ErrorValue::Num)),
        ("DAYS(1)", e(ErrorValue::Value)),
        ("DAYS(1/0, 1)", e(ErrorValue::Div0)),
    ]);
}

/// text compares by excel's collation, not by code point: punctuation and
/// symbols rank below digits, and digits below letters.
#[test]
fn text_compares_by_collation() {
    check(&[
        ("\"[a]\"<\"[a0]\"", b(true)),
        ("\"[b]\"<\"c\"", b(true)),
        ("\"_\"<\"1\"", b(true)),
        ("\"1\"<\"a\"", b(true)),
        ("\"a\"<\"b\"", b(true)),
        ("\"A\"=\"a\"", b(true)),
        ("\"apple\"<\"apples\"", b(true)),
        ("MATCH(\"[a]\",{\"[a]\";\"[a0]\"},0)", n(1.0)),
    ]);
}
