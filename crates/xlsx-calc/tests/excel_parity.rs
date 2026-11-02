//! results at numeric limits checked against workbooks recalculated in
//! Excel 365 for Mac. the formula each test cites is Excel's cached value;
//! the rows around it follow the rule the test states.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, Workbook};

fn recalculated(formulas: &[&str]) -> Vec<CellValue> {
    let mut book = Workbook {
        sheets: vec![Sheet::new("Sheet1")],
        ..Workbook::default()
    };
    for (row, formula) in formulas.iter().enumerate() {
        book.sheets[0].set_cell(
            CellRef::new(row as u32, 0),
            Cell {
                formula: Some((*formula).into()),
                ..Cell::default()
            },
        );
    }
    rebuild_and_recalc_all(&mut book, None);
    (0..formulas.len())
        .map(|row| {
            book.sheets[0]
                .cell(CellRef::new(row as u32, 0))
                .map(|cell| cell.value.clone())
                .unwrap_or_default()
        })
        .collect()
}

fn check(cases: &[(&str, CellValue)]) {
    let formulas: Vec<&str> = cases.iter().map(|(formula, _)| *formula).collect();
    for ((formula, want), got) in cases.iter().zip(recalculated(&formulas)) {
        assert_eq!(&got, want, "{formula}");
    }
}

fn n(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn t(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn e(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

/// Excel 365 for Mac result: MOD(2^53,7), MOD(1125900000000,1) and
/// MOD(2^49,7) are #NUM!, MOD(1125899999999,1) is 0 and MOD(2^40,7) is 2:
/// excel refuses a quotient of 1125900000000 or more.
#[test]
fn mod_refuses_a_quotient_past_its_limit() {
    check(&[
        ("MOD(2^53,7)", e(ErrorValue::Num)),
        ("MOD(1125900000000,1)", e(ErrorValue::Num)),
        ("MOD(-1125900000000,1)", e(ErrorValue::Num)),
        ("MOD(2251800000000,2)", e(ErrorValue::Num)),
        ("MOD(1125899999999,1)", n(0.0)),
        ("MOD(1125899999999,2)", n(1.0)),
        ("MOD(-1125899999999,2)", n(1.0)),
        ("MOD(1125900000001,1)", e(ErrorValue::Num)),
        ("MOD(-1125900000001,1)", e(ErrorValue::Num)),
        ("MOD(2^49,7)", e(ErrorValue::Num)),
        ("MOD(2^40,7)", n(2.0)),
        ("MOD(2^53,2^20)", n(0.0)),
        ("MOD(134217728,1)", n(0.0)),
        ("MOD(1E15,1E4)", n(0.0)),
    ]);
}

/// Excel 365 for Mac result: ATANH(0.9999999999999999) is 17.616361586450413,
/// the value at 0.999999999999999: a literal keeps 15 significant digits.
#[test]
fn number_literals_keep_fifteen_significant_digits() {
    check(&[
        ("ATANH(0.9999999999999999)", n(17.616_361_586_450_413)),
        ("ATANH(0.999999999999999)", n(17.616_361_586_450_413)),
        ("9007199254740993-9007199254740990", n(0.0)),
        ("9.99999999999999E+307", n(9.999_999_999_999_99e307)),
        ("0.1234567890123459", n(0.123_456_789_012_345)),
        (
            "0.1234567890123459=0.123456789012345",
            CellValue::Bool { value: true },
        ),
        ("1E-309*1", n(0.0)),
    ]);
}

/// Excel 365 for Mac result: AVERAGE and SUM of two 9.99999999999999E+307
/// and 9.99999999999999E+307*10 are #NUM!; AVERAGE of it and its negation
/// is 0.
#[test]
fn overflowing_results_are_num() {
    check(&[
        (
            "AVERAGE(9.99999999999999E+307,9.99999999999999E+307)",
            e(ErrorValue::Num),
        ),
        (
            "SUM(9.99999999999999E+307,9.99999999999999E+307)",
            e(ErrorValue::Num),
        ),
        ("9.99999999999999E+307*10", e(ErrorValue::Num)),
        (
            "AVERAGE(9.99999999999999E+307,-9.99999999999999E+307)",
            n(0.0),
        ),
    ]);
}

/// Excel 365 for Mac result: 1E-307/1E10, 1E-200*1E-110, EXP(-740),
/// POWER(10,-310), PRODUCT(1E-200,1E-110) and 2.2250738585072E-308*0.5 are 0:
/// a result below the smallest normal double is stored as 0. the math that
/// produces it keeps every bit, and normal results are untouched.
#[test]
fn subnormal_results_are_stored_as_zero() {
    check(&[
        ("1E-307/1E10", n(0.0)),
        ("-1E-307/1E10", n(0.0)),
        ("1E-200*1E-110", n(0.0)),
        ("EXP(-740)", n(0.0)),
        ("POWER(10,-310)", n(0.0)),
        ("PRODUCT(1E-200,1E-110)", n(0.0)),
        ("2.2250738585072E-308*0.5", n(0.0)),
        ("2.2250738585072E-308*1", n(0.0)),
        ("1E-300*1E-7", n(1e-307)),
        ("6E-308/2", n(3e-308)),
        ("0.1+0.2", n(0.1 + 0.2)),
    ]);
}

/// Excel 365 for Mac result: numbers become text with 15 significant
/// digits, plain unless the plain form (sign excluded) is longer than 20
/// characters, then scientific with a signed two-digit-or-more exponent.
#[test]
fn numbers_become_text_with_fifteen_significant_digits() {
    check(&[
        ("\"\"&(0.1+0.2)", t("0.3")),
        ("\"\"&(0.1+0.7)", t("0.8")),
        ("\"\"&(1/3)", t("0.333333333333333")),
        ("\"\"&1000000000000000", t("1000000000000000")),
        ("\"\"&123456789012345678", t("123456789012345000")),
        (
            "\"\"&(-2/3*100000000000000000000)",
            t("-66666666666666700000"),
        ),
        ("\"\"&(10^19)", t("10000000000000000000")),
        ("\"\"&(99999999999999.9*10^6)", t("99999999999999900000")),
        ("\"\"&(10^20)", t("1E+20")),
        ("\"\"&(1.5*10^20)", t("1.5E+20")),
        ("\"\"&(-1*10^21)", t("-1E+21")),
        ("\"\"&(10^100)", t("1E+100")),
        ("\"\"&(1.23456*10^200)", t("1.23456E+200")),
        ("\"\"&0.00001", t("0.00001")),
        ("\"\"&(10^-12)", t("0.000000000001")),
        ("\"\"&(1.5*10^-9)", t("0.0000000015")),
        ("\"\"&(1.23456789*10^-7)", t("0.000000123456789")),
        (
            "\"\"&(123456789012345678*10^-21)",
            t("0.000123456789012345"),
        ),
        ("\"\"&(10^-18)", t("0.000000000000000001")),
        ("\"\"&(10^-19)", t("1E-19")),
        ("\"\"&(1.5*10^-17)", t("0.000000000000000015")),
        ("\"\"&(1.5*10^-18)", t("1.5E-18")),
        ("\"\"&(1.23*10^-15)", t("0.00000000000000123")),
        ("\"\"&(123456789012345*10^-20)", t("1.23456789012345E-06")),
        ("\"\"&(12345678901234*10^-20)", t("1.2345678901234E-07")),
        ("\"\"&(123456789012345678/10^13)", t("12345.6789012345")),
        ("\"\"&(2/3*10^-12)", t("6.66666666666667E-13")),
        ("\"\"&(10^-100)", t("1E-100")),
        ("\"\"&(-1*10^-20)", t("-1E-20")),
        ("\"\"&2^-41", t("4.54747350886464E-13")),
        ("\"\"&9.99999999999999E+307", t("9.99999999999999E+307")),
        ("\"\"&0", t("0")),
        ("VALUETOTEXT(2/3)", t("0.666666666666667")),
        ("VALUETOTEXT(10^21)", t("1E+21")),
        (
            "TEXTJOIN(\",\",TRUE,10^-12,10^22)",
            t("0.000000000001,1E+22"),
        ),
        ("ARRAYTOTEXT({1E-12,1E22},1)", t("{0.000000000001,1E+22}")),
        ("CONCAT(0.1+0.2,\"|\",2/3)", t("0.3|0.666666666666667")),
        ("LEN(0.1+0.2)", n(3.0)),
    ]);
}

/// the 15-digit text is for display only: values that print alike stay
/// distinct where Excel compares them as numbers.
#[test]
fn number_keys_keep_the_exact_value() {
    check(&[
        ("ROWS(UNIQUE({0;2}+2^53))", n(2.0)),
        ("ROWS(UNIQUE({0;1}*2^-52+1))", n(2.0)),
    ]);
}

/// Excel 365 for Mac results for the remaining boundary cases of the
/// recalculated workbooks.
#[test]
fn boundary_cases_match_excel() {
    check(&[
        ("T.DIST.RT(1E+200,5)", e(ErrorValue::Num)),
        ("T.DIST.RT(1.4E+154,5)", e(ErrorValue::Num)),
        ("T.DIST.RT(1.3E+154,5)", n(0.0)),
        ("T.DIST.RT(1E+100,5)", n(0.0)),
        ("F.INV.RT(0.5,1E10,1E10)", n(1.0)),
        ("F.INV.RT(0.5,1E9,1E9)", n(1.0)),
        ("F.DIST.RT(1,1E10,1E10)", n(0.5)),
        ("F.DIST.RT(1,1.1E10,1E10)", e(ErrorValue::Num)),
        ("MULTINOMIAL(10000000000000000,1)", e(ErrorValue::Num)),
        ("BESSELK(1000,1000)", n(0.0)),
        (
            "TBILLPRICE(DATE(9999,1,1),DATE(9999,12,31),0.05)",
            e(ErrorValue::Num),
        ),
    ]);
}

/// Excel 365 for Mac result: XIRR of these flows is 0.3733625352382659, which
/// is not the root but a point 1.7E-9 above it (its NPV is -1.0E-5). no
/// stopping rule on newton's iterates from 0.1 reproduces it, so the engine
/// returns the root, 4.6E-9 relative below Excel.
#[test]
fn xirr_stays_within_five_billionths_of_excel() {
    let excel = 0.373_362_535_238_265_9;
    let values =
        recalculated(&["XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})"]);
    let [CellValue::Number { value }] = values.as_slice() else {
        panic!("XIRR is not a number: {values:?}");
    };
    let gap = (*value - excel).abs() / excel;
    assert!(gap < 5e-9, "{value} vs {excel}: {gap:e}");
}
