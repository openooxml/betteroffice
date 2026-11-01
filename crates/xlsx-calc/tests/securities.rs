//! securities and day-count financial functions, checked against
//! microsoft's documented examples and an independent python
//! reimplementation of excel's documented formulas: calendar dates from
//! python's `datetime`, sums and roots in mpmath at 30 digits.

use xlsx_calc::{EvalContext, Value, evaluate_array, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{Cell, CellRef, CellValue, DateSystem, ErrorValue, Sheet, SheetId, Workbook};

const RESULT: &str = "ZZ1";

fn eval_in(system: DateSystem, formula: &str) -> CellValue {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::parse_a1(RESULT).unwrap(),
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    let mut book = Workbook {
        sheets: vec![sheet],
        date_system: system,
        ..Workbook::default()
    };
    rebuild_and_recalc_all(&mut book, None);
    book.sheets[0]
        .cell(CellRef::parse_a1(RESULT).unwrap())
        .unwrap()
        .value
        .clone()
}

fn eval(formula: &str) -> CellValue {
    eval_in(DateSystem::V1900, formula)
}

fn value(formula: &str) -> f64 {
    match eval(formula) {
        CellValue::Number { value } => value,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

/// within `rel` of `want`, relative to its magnitude.
fn close(formula: &str, want: f64, rel: f64) {
    let got = value(formula);
    let scale = want.abs().max(f64::MIN_POSITIVE);
    assert!(
        (got - want).abs() <= rel * scale,
        "{formula}: {got:e} vs {want:e} (rel {:e})",
        (got - want).abs() / scale
    );
}

/// within `abs` of `want`: microsoft's examples print rounded results.
fn near(formula: &str, want: f64, abs: f64) {
    let got = value(formula);
    assert!((got - want).abs() <= abs, "{formula}: {got:e} vs {want:e}");
}

fn exact(formula: &str, want: f64) {
    assert_eq!(value(formula), want, "{formula}");
}

fn fails(formula: &str, want: ErrorValue) {
    assert_eq!(eval(formula), CellValue::Error { value: want }, "{formula}");
}

fn same(a: &str, b: &str) {
    assert_eq!(eval(a), eval(b), "{a} vs {b}");
}

#[test]
fn coupon_functions_microsoft_examples() {
    let args = "DATE(2011,1,25),DATE(2011,11,15),2,1";
    exact(&format!("COUPDAYBS({args})"), 71.0);
    exact(&format!("COUPDAYS({args})"), 181.0);
    exact(&format!("COUPDAYSNC({args})"), 110.0);
    exact(&format!("COUPNCD({args})"), 40_678.0);
    exact(&format!("COUPPCD({args})"), 40_497.0);
    exact("COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)", 4.0);
}

/// a maturity on the month's last day puts every coupon on a month end; the
/// 30/360 bases take COUPDAYSNC as the nominal period less COUPDAYBS, the
/// others count actual days.
#[test]
fn coupon_functions_match_the_python_reimplementation() {
    // python datetime reimplementation
    let cases: &[(&str, [f64; 6])] = &[
        (
            "DATE(2012,3,15),DATE(2016,8,31),4,0",
            [15.0, 90.0, 75.0, 40_968.0, 41_060.0, 18.0],
        ),
        (
            "DATE(2010,7,20),DATE(2020,2,29),2,4",
            [142.0, 180.0, 38.0, 40_237.0, 40_421.0, 20.0],
        ),
        (
            "DATE(2010,7,20),DATE(2020,2,29),2,3",
            [142.0, 182.5, 42.0, 40_237.0, 40_421.0, 20.0],
        ),
        (
            "DATE(2019,12,31),DATE(2025,6,30),1,1",
            [184.0, 366.0, 182.0, 43_646.0, 44_012.0, 6.0],
        ),
        (
            "DATE(2011,5,15),DATE(2011,11,15),2,2",
            [0.0, 180.0, 184.0, 40_678.0, 40_862.0, 1.0],
        ),
    ];
    for (args, [bs, days, nc, pcd, ncd, count]) in cases {
        exact(&format!("COUPDAYBS({args})"), *bs);
        exact(&format!("COUPDAYS({args})"), *days);
        exact(&format!("COUPDAYSNC({args})"), *nc);
        exact(&format!("COUPPCD({args})"), *pcd);
        exact(&format!("COUPNCD({args})"), *ncd);
        exact(&format!("COUPNUM({args})"), *count);
    }
}

/// excel truncates the dates, frequency and basis it is given.
#[test]
fn coupon_functions_truncate_their_arguments() {
    exact(
        "COUPDAYBS(DATE(2011,1,25)+0.75,DATE(2011,11,15)+0.2,2.9,1.6)",
        71.0,
    );
    exact("COUPNUM(\"1/25/2011\",\"11/15/2011\",2)", 2.0);
}

#[test]
fn coupon_functions_error_cases() {
    for name in [
        "COUPDAYBS",
        "COUPDAYS",
        "COUPDAYSNC",
        "COUPNCD",
        "COUPNUM",
        "COUPPCD",
    ] {
        fails(
            &format!("{name}(DATE(2011,11,15),DATE(2011,11,15),2,1)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2012,1,1),DATE(2011,11,15),2,1)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2011,1,25),DATE(2011,11,15),3,1)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2011,1,25),DATE(2011,11,15),2,5)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2011,1,25),DATE(2011,11,15),2,-1)"),
            ErrorValue::Num,
        );
        fails(&format!("{name}(-1,DATE(2011,11,15),2)"), ErrorValue::Num);
        fails(
            &format!("{name}(\"soon\",DATE(2011,11,15),2)"),
            ErrorValue::Value,
        );
        fails(
            &format!("{name}(DATE(2011,1,25),DATE(2011,11,15))"),
            ErrorValue::Value,
        );
    }
}

/// a 1904 workbook reads and writes coupon dates in its own serials.
#[test]
fn coupon_dates_follow_the_1904_date_system() {
    let formula = "COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)";
    assert_eq!(
        eval_in(DateSystem::V1904, formula),
        CellValue::Number {
            value: 40_678.0 - 1_462.0
        }
    );
    assert_eq!(
        eval_in(
            DateSystem::V1904,
            "COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)"
        ),
        CellValue::Number { value: 71.0 }
    );
}

#[test]
fn accrint_microsoft_examples() {
    let tail = "DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0";
    near(&format!("ACCRINT(DATE(2008,3,1),{tail})"), 16.666_667, 5e-7);
    near(
        &format!("ACCRINT(DATE(2008,3,5),{tail},FALSE)"),
        15.555_556,
        5e-7,
    );
    near(
        &format!("ACCRINT(DATE(2008,4,5),{tail},TRUE)"),
        7.222_222_2,
        5e-8,
    );
}

/// the sum over the first interest date's quasi-coupon periods: one coupon
/// for a whole period, days over the period's length for a part.
#[test]
fn accrint_matches_the_python_reimplementation() {
    // mpmath 30 digits over python datetime
    close(
        "ACCRINT(DATE(2007,9,15),DATE(2008,3,15),DATE(2009,1,10),0.06,1000,2,0)",
        79.166_666_666_666_66,
        1e-12,
    );
    close(
        "ACCRINT(DATE(2007,9,15),DATE(2008,3,15),DATE(2009,1,10),0.06,1000,2,0,FALSE)",
        49.166_666_666_666_664,
        1e-12,
    );
    close(
        "ACCRINT(DATE(2008,1,10),DATE(2008,9,30),DATE(2008,11,20),0.05,500,4,1)",
        21.527_860_726_230_294,
        1e-12,
    );
    close(
        "ACCRINT(DATE(2010,2,1),DATE(2010,8,31),DATE(2010,6,1),0.07,1000,2,3)",
        23.013_698_630_136_99,
        1e-12,
    );
}

#[test]
fn accrint_defaults_an_omitted_par_to_1000() {
    close(
        "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,,2)",
        16.666_666_666_666_668,
        1e-12,
    );
}

#[test]
fn accrint_error_cases() {
    let first = "DATE(2008,8,31)";
    fails(
        &format!("ACCRINT(DATE(2008,5,1),{first},DATE(2008,5,1),0.1,1000,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ACCRINT(DATE(2008,3,1),{first},DATE(2008,5,1),0,1000,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ACCRINT(DATE(2008,3,1),{first},DATE(2008,5,1),0.1,-5,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ACCRINT(DATE(2008,3,1),{first},DATE(2008,5,1),0.1,1000,12)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ACCRINT(DATE(2008,3,1),{first},DATE(2008,5,1),0.1,1000,2,7)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ACCRINT(DATE(2008,3,1),{first},DATE(2008,5,1),\"x\",1000,2)"),
        ErrorValue::Value,
    );
}

#[test]
fn accrintm_microsoft_example_and_oracles() {
    near(
        "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)",
        20.547_945_21,
        5e-9,
    );
    // mpmath 30 digits over python datetime
    close(
        "ACCRINTM(DATE(2007,11,30),DATE(2009,2,28),0.045,2500,1)",
        140.419_708_029_197_07,
        1e-12,
    );
    close(
        "ACCRINTM(DATE(2010,1,31),DATE(2010,7,31),0.06)",
        30.0,
        1e-12,
    );
    fails(
        "ACCRINTM(DATE(2008,6,15),DATE(2008,6,15),0.1)",
        ErrorValue::Num,
    );
    fails(
        "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0)",
        ErrorValue::Num,
    );
    fails(
        "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,0)",
        ErrorValue::Num,
    );
    fails(
        "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,5)",
        ErrorValue::Num,
    );
}

#[test]
fn price_microsoft_example() {
    near(
        "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
        94.634_36,
        5e-6,
    );
}

/// excel's documented sum over the coupons left, with a single coupon left
/// too.
#[test]
fn price_matches_mpmath() {
    // mpmath 30 digits
    close(
        "PRICE(DATE(2010,3,1),DATE(2025,8,31),0.045,0.052,100,2,1)",
        92.613_877_050_914_91,
        1e-12,
    );
    close(
        "PRICE(DATE(2015,6,30),DATE(2017,2,28),0.08,0.03,105,4,3)",
        112.852_511_152_956_15,
        1e-12,
    );
    close(
        "PRICE(DATE(2011,1,25),DATE(2011,6,15),0.06,0.05,100,2,0)",
        100.374_055_250_478_58,
        1e-12,
    );
    close(
        "PRICE(DATE(2009,10,10),DATE(2019,4,1),0.07,0.065,100,1,2)",
        103.305_888_848_930_08,
        1e-12,
    );
    close(
        "PRICE(DATE(2013,1,31),DATE(2023,7,31),0.05,0.04,100,2,4)",
        108.505_604_580_684_86,
        1e-12,
    );
}

#[test]
fn price_error_cases() {
    let dates = "DATE(2008,2,15),DATE(2017,11,15)";
    fails(
        &format!("PRICE({dates},-0.01,0.065,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("PRICE({dates},0.0575,-0.01,100,2)"),
        ErrorValue::Num,
    );
    fails(&format!("PRICE({dates},0.0575,0.065,0,2)"), ErrorValue::Num);
    fails(
        &format!("PRICE({dates},0.0575,0.065,100,3)"),
        ErrorValue::Num,
    );
    fails(
        &format!("PRICE({dates},0.0575,0.065,100,2,5)"),
        ErrorValue::Num,
    );
    fails(
        "PRICE(DATE(2017,11,15),DATE(2017,11,15),0.0575,0.065,100,2)",
        ErrorValue::Num,
    );
    fails(
        &format!("PRICE({dates},0.0575,0.065,100)"),
        ErrorValue::Value,
    );
}

#[test]
fn yield_microsoft_example() {
    near(
        "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
        0.065,
        1e-8,
    );
}

/// several coupons left: an iterative root of PRICE, held to 1e-9. one
/// coupon left: excel's closed form, exact.
#[test]
fn yield_matches_mpmath() {
    // mpmath 30 digits, findroot on the price sum
    close(
        "YIELD(DATE(2010,3,1),DATE(2025,8,31),0.045,93.5,100,2,1)",
        0.051_123_495_160_106_48,
        1e-9,
    );
    close(
        "YIELD(DATE(2009,10,10),DATE(2019,4,1),0.07,110,100,1,2)",
        0.055_923_412_391_249_95,
        1e-9,
    );
    close(
        "YIELD(DATE(2013,1,31),DATE(2023,7,31),0,70,100,4,4)",
        0.034_113_688_426_713_27,
        1e-9,
    );
    // mpmath 30 digits, closed form
    close(
        "YIELD(DATE(2011,1,25),DATE(2011,6,15),0.06,100.2,100,2,0)",
        0.054_385_799_263_525_56,
        1e-12,
    );
    close(
        "YIELD(DATE(2011,1,25),DATE(2011,6,15),0.06,100.2,100,2,1)",
        0.054_360_774_626_281_76,
        1e-12,
    );
}

#[test]
fn yield_inverts_price() {
    close(
        "YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,\
         PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0),100,2,0)",
        0.065,
        1e-12,
    );
}

#[test]
fn yield_error_cases() {
    let dates = "DATE(2008,2,15),DATE(2016,11,15)";
    fails(&format!("YIELD({dates},-0.01,95,100,2)"), ErrorValue::Num);
    fails(&format!("YIELD({dates},0.0575,0,100,2)"), ErrorValue::Num);
    fails(&format!("YIELD({dates},0.0575,95,0,2)"), ErrorValue::Num);
    fails(&format!("YIELD({dates},0.0575,95,100,6)"), ErrorValue::Num);
    fails(
        &format!("YIELD({dates},0.0575,95,100,2,9)"),
        ErrorValue::Num,
    );
    fails(
        "YIELD(DATE(2016,11,15),DATE(2008,2,15),0.0575,95,100,2)",
        ErrorValue::Num,
    );
}

#[test]
fn duration_microsoft_examples() {
    near(
        "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)",
        10.919_145_3,
        5e-8,
    );
    near(
        "MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
        5.735_67,
        5e-6,
    );
}

/// payments are timed by COUPDAYSNC over COUPDAYS into their period, which
/// is what reproduces microsoft's example; a zero coupon's duration is its
/// term.
#[test]
fn duration_matches_mpmath() {
    // mpmath 30 digits
    close(
        "DURATION(DATE(2010,3,1),DATE(2025,8,31),0.045,0.052,2,1)",
        11.130_758_893_202_499,
        1e-12,
    );
    close(
        "DURATION(DATE(2015,6,30),DATE(2017,2,28),0.08,0.03,4)",
        1.571_789_917_864_943,
        1e-12,
    );
    close(
        "DURATION(DATE(2013,1,31),DATE(2023,7,31),0,0.04,2,4)",
        10.5,
        1e-12,
    );
    close(
        "MDURATION(DATE(2010,3,1),DATE(2025,8,31),0.045,0.052,2,1)",
        10.848_692_878_365_009,
        1e-12,
    );
    close(
        "MDURATION(DATE(2015,6,30),DATE(2017,2,28),0.08,0.03,4)",
        1.560_089_248_501_184,
        1e-12,
    );
}

#[test]
fn duration_error_cases() {
    for name in ["DURATION", "MDURATION"] {
        let dates = "DATE(2018,7,1),DATE(2048,1,1)";
        fails(&format!("{name}({dates},-0.08,0.09,2)"), ErrorValue::Num);
        fails(&format!("{name}({dates},0.08,-0.09,2)"), ErrorValue::Num);
        fails(&format!("{name}({dates},0.08,0.09,5)"), ErrorValue::Num);
        fails(&format!("{name}({dates},0.08,0.09,2,5)"), ErrorValue::Num);
        fails(
            &format!("{name}(DATE(2048,1,1),DATE(2048,1,1),0.08,0.09,2)"),
            ErrorValue::Num,
        );
    }
}

#[test]
fn discount_securities_microsoft_examples() {
    near(
        "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)",
        99.795_83,
        5e-6,
    );
    near(
        "DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)",
        0.052_420_213,
        5e-10,
    );
    near(
        "INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)",
        0.057_68,
        5e-6,
    );
    near(
        "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)",
        1_014_584.654,
        5e-4,
    );
    near(
        "YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)",
        0.052_823,
        5e-7,
    );
}

/// each is a closed form over YEARFRAC of settlement to maturity.
#[test]
fn discount_securities_match_mpmath() {
    // mpmath 30 digits over python datetime
    close(
        "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)",
        99.795_833_333_333_33,
        1e-12,
    );
    close(
        "PRICEDISC(DATE(2019,5,31),DATE(2020,2,29),0.04,105,1)",
        101.855_737_704_918_03,
        1e-12,
    );
    close(
        "DISC(DATE(2020,1,31),DATE(2020,12,31),96.5,100,4)",
        0.038_181_818_181_818_185,
        1e-12,
    );
    close(
        "INTRATE(DATE(2015,8,31),DATE(2016,2,29),10000,10500)",
        0.100_558_659_217_877_09,
        1e-12,
    );
    close(
        "RECEIVED(DATE(2011,3,31),DATE(2012,3,31),5000,0.03,3)",
        5_155.075_984_407_661,
        1e-12,
    );
    close(
        "YIELDDISC(DATE(2016,1,1),DATE(2016,12,31),95,100,1)",
        0.052_775_775_054_073_54,
        1e-12,
    );
}

#[test]
fn discount_securities_error_cases() {
    for name in ["PRICEDISC", "DISC", "INTRATE", "RECEIVED", "YIELDDISC"] {
        let dates = "DATE(2008,2,15),DATE(2008,5,15)";
        fails(&format!("{name}({dates},0,100)"), ErrorValue::Num);
        fails(&format!("{name}({dates},0.05,-1)"), ErrorValue::Num);
        fails(&format!("{name}({dates},0.05,100,5)"), ErrorValue::Num);
        fails(
            &format!("{name}(DATE(2008,5,15),DATE(2008,5,15),0.05,100)"),
            ErrorValue::Num,
        );
        fails(&format!("{name}({dates},\"x\",100)"), ErrorValue::Value);
    }
}

#[test]
fn interest_at_maturity_microsoft_examples() {
    near(
        "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
        99.984_498_88,
        5e-9,
    );
    near(
        "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)",
        0.060_954,
        5e-7,
    );
}

#[test]
fn interest_at_maturity_matches_mpmath() {
    // mpmath 30 digits over python datetime
    close(
        "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
        99.984_498_875_556_95,
        1e-12,
    );
    close(
        "PRICEMAT(DATE(2012,7,1),DATE(2014,1,31),DATE(2011,12,31),0.05,0.045,1)",
        100.579_727_127_682,
        1e-12,
    );
    close(
        "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)",
        0.060_954_333_691_538_52,
        1e-12,
    );
    close(
        "YIELDMAT(DATE(2012,7,1),DATE(2014,1,31),DATE(2011,12,31),0.05,101,3)",
        0.042_215_590_487_971_9,
        1e-12,
    );
}

#[test]
fn interest_at_maturity_error_cases() {
    let dates = "DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11)";
    fails(&format!("PRICEMAT({dates},-0.01,0.061)"), ErrorValue::Num);
    fails(&format!("PRICEMAT({dates},0.061,-0.01)"), ErrorValue::Num);
    fails(&format!("YIELDMAT({dates},-0.01,100)"), ErrorValue::Num);
    fails(&format!("YIELDMAT({dates},0.061,0)"), ErrorValue::Num);
    fails(&format!("PRICEMAT({dates},0.061,0.061,5)"), ErrorValue::Num);
    fails(
        "YIELDMAT(DATE(2008,4,13),DATE(2008,4,13),DATE(2007,11,11),0.061,100)",
        ErrorValue::Num,
    );
}

#[test]
fn treasury_bills_microsoft_examples() {
    near(
        "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)",
        98.45,
        5e-12,
    );
    near(
        "TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)",
        0.0914,
        5e-5,
    );
    near(
        "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)",
        0.094_151,
        5e-7,
    );
}

/// a bill past 182 days gets the bond-equivalent quadratic, as excel gives.
#[test]
fn treasury_bills_match_mpmath() {
    // mpmath 30 digits
    close(
        "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)",
        0.094_151_493_565_943_02,
        1e-12,
    );
    close(
        "TBILLEQ(DATE(2020,1,15),DATE(2020,7,15),0.05)",
        0.052_009_119_407_238_53,
        1e-12,
    );
    close(
        "TBILLEQ(DATE(2020,1,15),DATE(2020,12,31),0.05)",
        0.052_627_651_808_428_565,
        1e-12,
    );
    close(
        "TBILLPRICE(DATE(2020,1,15),DATE(2020,12,31),0.05)",
        95.125,
        1e-12,
    );
    close(
        "TBILLYIELD(DATE(2020,1,15),DATE(2020,12,31),96)",
        0.042_735_042_735_042_736,
        1e-12,
    );
}

/// a bill may run one calendar year and no more.
#[test]
fn treasury_bills_error_cases() {
    for name in ["TBILLPRICE", "TBILLYIELD", "TBILLEQ"] {
        assert!(matches!(
            eval(&format!("{name}(DATE(2020,1,15),DATE(2021,1,15),0.05)")),
            CellValue::Number { .. }
        ));
        fails(
            &format!("{name}(DATE(2020,1,15),DATE(2021,1,16),0.05)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2020,1,15),DATE(2020,1,15),0.05)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2020,1,15),DATE(2020,6,1),0)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(DATE(2020,1,15),DATE(2020,6,1))"),
            ErrorValue::Value,
        );
    }
}

#[test]
fn odd_last_period_microsoft_examples() {
    near(
        "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
        99.88,
        5e-3,
    );
    near(
        "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)",
        0.045_192,
        5e-7,
    );
}

/// excel's documented sums over the quasi-coupon periods after
/// last_interest, a whole period counting one coupon.
#[test]
fn odd_last_period_matches_mpmath() {
    // mpmath 30 digits over python datetime
    close(
        "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
        99.878_286_014_721_34,
        1e-12,
    );
    close(
        "ODDLPRICE(DATE(2010,3,1),DATE(2010,11,30),DATE(2009,12,31),0.06,0.05,100,4,1)",
        100.685_947_526_956_82,
        1e-12,
    );
    close(
        "ODDLPRICE(DATE(2012,5,10),DATE(2012,9,20),DATE(2012,2,20),0.05,0.07,100,2,3)",
        99.268_780_583_671_22,
        1e-12,
    );
    close(
        "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)",
        0.045_192_235_629_168_85,
        1e-12,
    );
    close(
        "ODDLYIELD(DATE(2010,3,1),DATE(2010,11,30),DATE(2009,12,31),0.06,101,100,4,1)",
        0.045_735_827_583_754_18,
        1e-12,
    );
    close(
        "ODDLYIELD(DATE(2012,5,10),DATE(2012,9,20),DATE(2012,2,20),0.05,99,100,2,3)",
        0.077_557_209_020_910_01,
        1e-12,
    );
}

#[test]
fn odd_last_period_error_cases() {
    let dates = "DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15)";
    fails(
        &format!("ODDLPRICE({dates},-0.01,0.04,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDLPRICE({dates},0.0375,-0.01,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDLYIELD({dates},0.0375,0,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDLYIELD({dates},-0.01,99,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDLPRICE({dates},0.0375,0.04,100,3)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDLPRICE({dates},0.0375,0.04,100,2,5)"),
        ErrorValue::Num,
    );
    fails(
        "ODDLPRICE(DATE(2007,10,15),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.04,100,2)",
        ErrorValue::Num,
    );
    fails(
        "ODDLYIELD(DATE(2008,6,15),DATE(2008,6,15),DATE(2007,10,15),0.0375,99,100,2)",
        ErrorValue::Num,
    );
}

#[test]
fn odd_first_period_microsoft_examples() {
    near(
        "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
        113.60,
        5e-3,
    );
    near(
        "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)",
        0.0772,
        5e-5,
    );
}

/// short and long first periods by excel's documented formulas; the long
/// one sums the quasi-coupon periods back to issue and counts the whole
/// periods between settlement and the first coupon.
#[test]
fn odd_first_period_matches_mpmath() {
    // mpmath 30 digits over python datetime
    close(
        "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)",
        113.597_717_474_078_83,
        1e-12,
    );
    close(
        "ODDFPRICE(DATE(2010,2,10),DATE(2015,6,15),DATE(2010,1,20),DATE(2010,6,15),0.05,0.06,100,2,0)",
        95.484_978_918_800_41,
        1e-12,
    );
    close(
        "ODDFPRICE(DATE(2010,2,10),DATE(2015,6,15),DATE(2009,9,1),DATE(2010,6,15),0.05,0.06,100,2,0)",
        95.445_754_465_553_26,
        1e-12,
    );
    close(
        "ODDFPRICE(DATE(2009,11,20),DATE(2014,12,31),DATE(2009,5,5),DATE(2010,12,31),0.045,0.04,100,2,1)",
        102.126_903_985_207_58,
        1e-12,
    );
    // mpmath 30 digits, findroot on the price
    close(
        "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)",
        0.077_245_541_597_817_4,
        1e-9,
    );
    close(
        "ODDFYIELD(DATE(2010,2,10),DATE(2015,6,15),DATE(2010,1,20),DATE(2010,6,15),0.05,97,100,2,0)",
        0.056_584_556_991_064_73,
        1e-9,
    );
    close(
        "ODDFYIELD(DATE(2009,11,20),DATE(2014,12,31),DATE(2009,5,5),DATE(2010,12,31),0.045,102,100,2,1)",
        0.040_271_228_581_388_41,
        1e-9,
    );
}

#[test]
fn odd_first_period_error_cases() {
    let dates = "DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1)";
    fails(
        &format!("ODDFPRICE({dates},-0.01,0.06,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDFPRICE({dates},0.07,-0.01,100,2)"),
        ErrorValue::Num,
    );
    fails(&format!("ODDFYIELD({dates},0.07,0,100,2)"), ErrorValue::Num);
    fails(
        &format!("ODDFYIELD({dates},-0.01,90,100,2)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDFPRICE({dates},0.07,0.06,100,3)"),
        ErrorValue::Num,
    );
    fails(
        &format!("ODDFPRICE({dates},0.07,0.06,100,2,5)"),
        ErrorValue::Num,
    );
    fails(
        "ODDFPRICE(DATE(2008,10,15),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.07,0.06,100,2)",
        ErrorValue::Num,
    );
    fails(
        "ODDFPRICE(DATE(2009,3,1),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.07,0.06,100,2)",
        ErrorValue::Num,
    );
    fails(
        "ODDFYIELD(DATE(2008,11,11),DATE(2009,3,1),DATE(2008,10,15),DATE(2009,3,1),0.07,90,100,2)",
        ErrorValue::Num,
    );
}

#[test]
fn french_depreciation_microsoft_examples() {
    exact(
        "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
        776.0,
    );
    exact(
        "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
        360.0,
    );
}

/// whole schedules from the python reimplementation: AMORLINC prorates the
/// first period and stops at salvage; AMORDEGRC rounds every charge and
/// halves the remaining value in the period that would cut into salvage.
#[test]
fn french_depreciation_schedules() {
    // python reimplementation of the documented schedules
    let linc = [
        131.803_278_688_524_59,
        360.0,
        360.0,
        360.0,
        360.0,
        360.0,
        168.196_721_311_475_41,
        0.0,
        0.0,
    ];
    for (period, want) in linc.iter().enumerate() {
        let formula =
            format!("AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,{period},0.15,1)");
        if *want == 0.0 {
            exact(&formula, 0.0);
        } else {
            close(&formula, *want, 1e-12);
        }
    }
    let degrc = [330.0, 776.0, 485.0, 303.0, 190.0, 158.0, 0.0, 0.0];
    for (period, want) in degrc.iter().enumerate() {
        exact(
            &format!("AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,{period},0.15,1)"),
            *want,
        );
    }
    for (period, want) in [3750.0, 2812.0, 1547.0, 851.0, 468.0, 286.0]
        .iter()
        .enumerate()
    {
        exact(
            &format!("AMORDEGRC(10000,DATE(2015,3,1),DATE(2015,12,31),500,{period},0.3,0)"),
            *want,
        );
    }
    close(
        "AMORLINC(10000,DATE(2015,3,1),DATE(2015,12,31),500,0,0.3,3)",
        2_506.849_315_068_493,
        1e-12,
    );
    close(
        "AMORLINC(10000,DATE(2015,3,1),DATE(2015,12,31),500,3,0.3,3)",
        993.150_684_931_506_9,
        1e-12,
    );
}

/// excel documents `#NUM!` for a life between 0 and 1, 1 and 2, 2 and 3 or
/// 4 and 5 years; both functions leave out basis 2.
#[test]
fn french_depreciation_error_cases() {
    let head = "2400,DATE(2008,8,19),DATE(2008,12,31),300,1";
    for rate in ["0.6", "0.4", "0.22"] {
        fails(&format!("AMORDEGRC({head},{rate},1)"), ErrorValue::Num);
    }
    assert!(matches!(
        eval(&format!("AMORDEGRC({head},0.25,1)")),
        CellValue::Number { .. }
    ));
    for name in ["AMORDEGRC", "AMORLINC"] {
        fails(&format!("{name}({head},0.15,2)"), ErrorValue::Num);
        fails(&format!("{name}({head},0.15,5)"), ErrorValue::Num);
        fails(&format!("{name}({head},0,1)"), ErrorValue::Num);
        fails(
            &format!("{name}(2400,DATE(2009,1,1),DATE(2008,12,31),300,1,0.15,1)"),
            ErrorValue::Num,
        );
        fails(
            &format!("{name}(2400,DATE(2008,8,19),DATE(2008,12,31),300,-1,0.15,1)"),
            ErrorValue::Num,
        );
        fails(&format!("{name}({head},\"x\",1)"), ErrorValue::Value);
    }
}

/// evaluate in array mode as (rows, cols, values).
fn arrayed(formula: &str) -> (usize, usize, Vec<CellValue>) {
    let book = Workbook {
        sheets: vec![Sheet::new("Sheet1")],
        ..Workbook::default()
    };
    let expression = parse_formula(formula).unwrap();
    let context = EvalContext::new(&book, SheetId(0));
    match evaluate_array(&expression, &context) {
        Value::Array(array) => {
            let mut values = Vec::new();
            for row in 0..array.rows() {
                for col in 0..array.cols() {
                    values.push(array.at(row, col));
                }
            }
            (array.rows(), array.cols(), values)
        }
        other => (1, 1, vec![other.into_scalar()]),
    }
}

/// an array formula's elements equal the scalar formulas they stand for.
fn lifts_to(formula: &str, rows: usize, cols: usize, scalars: &[&str]) {
    let (r, c, values) = arrayed(formula);
    assert_eq!((r, c), (rows, cols), "{formula}");
    let want: Vec<CellValue> = scalars.iter().map(|scalar| eval(scalar)).collect();
    assert_eq!(values, want, "{formula}");
}

#[test]
fn bond_functions_lift_over_an_array_of_yields() {
    lifts_to(
        "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,{0.06,0.065,0.07},100,2,0)",
        1,
        3,
        &[
            "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.06,100,2,0)",
            "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
            "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.07,100,2,0)",
        ],
    );
    lifts_to(
        "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,{95.04287;0},100,2,0)",
        2,
        1,
        &[
            "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)",
            "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,0,100,2,0)",
        ],
    );
}

/// a scalar settlement broadcasts against a column of maturities; a bad
/// frequency fails only its own element.
#[test]
fn coupon_functions_lift_with_per_element_errors() {
    lifts_to(
        "COUPNUM(40568,{40862;41228},2)",
        2,
        1,
        &["COUPNUM(40568,40862,2)", "COUPNUM(40568,41228,2)"],
    );
    let (_, _, values) = arrayed("COUPNUM(40568,{40862;41228},2)");
    assert_eq!(
        values,
        vec![
            CellValue::Number { value: 2.0 },
            CellValue::Number { value: 4.0 }
        ]
    );
    lifts_to(
        "COUPDAYS(40568,40862,{2,3,4},1)",
        1,
        3,
        &[
            "COUPDAYS(40568,40862,2,1)",
            "COUPDAYS(40568,40862,3,1)",
            "COUPDAYS(40568,40862,4,1)",
        ],
    );
    assert_eq!(
        arrayed("COUPDAYS(40568,40862,{2,3,4},1)").2[1],
        CellValue::Error {
            value: ErrorValue::Num
        }
    );
}

#[test]
fn bill_and_depreciation_functions_lift_over_blocks() {
    lifts_to(
        "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),{0.09,0.05;0.01,0})",
        2,
        2,
        &[
            "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)",
            "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.05)",
            "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.01)",
            "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0)",
        ],
    );
    lifts_to(
        "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,{0,1,2},0.15,1)",
        1,
        3,
        &[
            "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,1)",
            "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
            "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,2,0.15,1)",
        ],
    );
    lifts_to(
        "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,{0.0405,-1},100,2,0)",
        1,
        2,
        &[
            "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)",
            "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,-1,100,2,0)",
        ],
    );
}

/// a prorated first period worth more than cost less salvage takes only
/// that, and later periods take nothing; neither function depreciates below
/// salvage or charges a negative amount.
#[test]
fn french_depreciation_stops_at_salvage() {
    let head = "2400,DATE(2008,8,19),DATE(2008,12,31)";
    close(&format!("AMORLINC({head},2300,0,0.15,1)"), 100.0, 1e-12);
    for period in 1..4 {
        exact(&format!("AMORLINC({head},2300,{period},0.15,1)"), 0.0);
    }
    exact(&format!("AMORDEGRC({head},2200,0,0.15,1)"), 200.0);
    for period in 1..4 {
        exact(&format!("AMORDEGRC({head},2200,{period},0.15,1)"), 0.0);
    }
}

/// a schedule whose charges no longer move the remaining value runs out of
/// recalculation budget instead of looping forever.
#[test]
fn amordegrc_is_bounded_by_the_recalculation_budget() {
    fails(
        "AMORDEGRC(1E20,DATE(2008,8,19),DATE(2008,12,31),0,1E16,1E-18,1)",
        ErrorValue::Num,
    );
    fails(
        "AMORDEGRC(1E9,DATE(2008,8,19),DATE(2008,12,31),0,1E9,1E-9,1)",
        ErrorValue::Num,
    );
}

/// microsoft documents no redemption check for the odd-period functions;
/// they reject a redemption at or below zero as PRICE and YIELD do.
#[test]
fn odd_period_functions_reject_a_non_positive_redemption() {
    fails(
        "ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,0,2,1)",
        ErrorValue::Num,
    );
    fails(
        "ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,0,2,0)",
        ErrorValue::Num,
    );
    fails(
        "ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,0,2,0)",
        ErrorValue::Num,
    );
    fails(
        "ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,-1,2,0)",
        ErrorValue::Num,
    );
}

/// settlement must come after issue.
#[test]
fn interest_at_maturity_needs_settlement_after_issue() {
    fails(
        "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2008,3,15),0.061,0.061,0)",
        ErrorValue::Num,
    );
    fails(
        "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2008,6,15),0.0625,100.0123,0)",
        ErrorValue::Num,
    );
    fails(
        "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2008,2,15),0.061,0.061,0)",
        ErrorValue::Num,
    );
}

/// a discount that eats the whole face value is `#NUM!`, not a negative
/// price or redemption.
#[test]
fn discounts_past_the_face_value_are_num_errors() {
    fails(
        "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),6)",
        ErrorValue::Num,
    );
    fails("TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),6)", ErrorValue::Num);
    fails("TBILLEQ(DATE(2008,1,1),DATE(2008,7,2),3)", ErrorValue::Num);
    fails(
        "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,5,2)",
        ErrorValue::Num,
    );
    fails(
        "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,4,2)",
        ErrorValue::Num,
    );
}

/// microsoft scopes calc_method to a settlement after first_interest, so
/// before it both methods accrue from issue.
#[test]
fn accrint_calc_method_only_matters_after_first_interest() {
    same(
        "ACCRINT(DATE(2007,12,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE)",
        "ACCRINT(DATE(2007,12,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE)",
    );
    // python reimplementation, 30/360 split at the 2008-02-29 quasi-coupon date
    close(
        "ACCRINT(DATE(2007,12,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE)",
        41.388_888_888_888_89,
        1e-12,
    );
}
