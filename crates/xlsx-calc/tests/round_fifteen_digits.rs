//! ROUND takes its operand at 15 significant digits, as excel does, then rounds
//! half away from zero in decimal. expected values come from excel's cached
//! results where noted, otherwise from a python decimal oracle: the exact
//! binary value rounded half up to 15 significant digits, then half up at
//! `digits` decimals. literals past 15 significant digits are built by
//! arithmetic, since a literal keeps only 15.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, Sheet, Workbook};

fn eval(formula: &str) -> CellValue {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    let mut book = Workbook {
        sheets: vec![sheet],
        ..Workbook::default()
    };
    rebuild_and_recalc_all(&mut book, None);
    book.sheets[0]
        .cell(CellRef::parse_a1("A1").unwrap())
        .unwrap()
        .value
        .clone()
}

fn exact(cases: &[(&str, f64)]) {
    for (formula, want) in cases {
        match eval(formula) {
            CellValue::Number { value } => assert!(
                value.to_bits() == want.to_bits() || (value == 0.0 && *want == 0.0),
                "{formula}: {value:e} vs {want:e}"
            ),
            other => panic!("{formula}: expected {want:e}, got {other:?}"),
        }
    }
}

#[test]
fn round_of_a_sixteen_digit_operand_matches_excels_cache() {
    exact(&[
        ("ROUND(1000000000000000+0.75,0)", 1e15),
        ("FLOOR.MATH(1000000000000000+0.75)", 1e15),
        ("CEILING.MATH(-1000000000000000-0.75)", -1e15),
    ]);
}

#[test]
fn mround_and_int_keep_the_binary_operand_as_excel_does() {
    exact(&[
        ("MROUND(2^53,3)", 9_007_199_254_740_992.0),
        ("INT(2^53+1)", 9_007_199_254_740_992.0),
    ]);
}

#[test]
fn decimal_ties_round_away_from_zero() {
    exact(&[
        ("ROUND(2.675,2)", 2.68),
        ("ROUND(-2.675,2)", -2.68),
        ("ROUND(1.005,2)", 1.01),
        ("ROUND(-1.005,2)", -1.01),
        ("ROUND(0.285,2)", 0.29),
        ("ROUND(-0.285,2)", -0.29),
        ("ROUND(1.45,1)", 1.5),
        ("ROUND(-4.35,1)", -4.4),
        ("ROUND(8.345,2)", 8.35),
        ("ROUND(-8.345,2)", -8.35),
    ]);
}

#[test]
fn negative_operands_mirror_positive_ones() {
    exact(&[
        ("ROUND(-2.5,0)", -3.0),
        ("ROUND(-0.5,0)", -1.0),
        ("ROUND(-0.4,0)", 0.0),
        ("ROUND(-1234.5678,-2)", -1200.0),
        ("ROUND(-0.0000123456789012345,18)", -1.2345678901235e-5),
        ("ROUND(-123456789012344-0.5,0)", -123_456_789_012_345.0),
        ("ROUND(-(2-2^-52)*2^1023,-300)", -1.797_693_13e308),
        ("ROUND(-1E-200,200)", -1e-200),
        ("ROUND(-1E-200,199)", 0.0),
    ]);
}

#[test]
fn operands_within_an_ulp_of_a_tie_round_on_their_fifteen_digit_value() {
    exact(&[
        ("ROUND(0.5-2^-54,0)", 1.0),
        ("ROUND(2.5-2^-51,0)", 3.0),
        ("ROUND(2.5+2^-51,0)", 3.0),
        ("ROUND(320.5-4E-13,0)", 321.0),
        ("ROUND(-2.5+2^-51,0)", -3.0),
        ("ROUND(123456789012344+0.5,0)", 123_456_789_012_345.0),
        ("ROUND(1000000000000000+0.5,0)", 1e15),
    ]);
}

#[test]
fn tiny_operands_round_exactly() {
    exact(&[
        ("ROUND(1E-300,300)", 1e-300),
        ("ROUND(1E-300,301)", 1e-300),
        ("ROUND(1.5E-300,300)", 2e-300),
        ("ROUND(1.23456789012345E-300,310)", 1.234_567_890_1e-300),
        ("ROUND(2.5E-308,308)", 3e-308),
        ("ROUND(1E-300,-5)", 0.0),
    ]);
}

#[test]
fn huge_operands_round_exactly() {
    exact(&[
        ("ROUND(9.99999999999999E+307,-307)", 1e308),
        ("ROUND(5E+307,-308)", 1e308),
        ("ROUND(9E+307,-309)", 0.0),
        ("ROUND((2-2^-52)*2^1023,-300)", 1.797_693_13e308),
        ("ROUND(1E+200,-199)", 1e200),
        ("ROUND(1.5E+22,-22)", 2e22),
        ("ROUND(1E+19+2048,-3)", 1e19),
    ]);
}

/// an operand already exact at `digits` comes back unchanged, so integers of
/// 16 or more digits keep their last digit, as `MROUND(2^53,3)` and
/// `INT(2^53+1)` keep 2^53 in excel.
#[test]
fn operands_exact_at_digits_are_kept() {
    exact(&[
        ("ROUND(2^53,0)", 9_007_199_254_740_992.0),
        ("ROUND(2^53+2,0)", 9_007_199_254_740_994.0),
        ("ROUND(-(2^53),0)", -9_007_199_254_740_992.0),
        ("ROUND(1E+16+2,0)", 10_000_000_000_000_002.0),
        ("ROUND(1000000000000000+0.75,2)", 1e15 + 0.75),
        ("ROUND((2-2^-52)*2^1023,0)", f64::MAX),
        ("ROUND((2-2^-52)*2^1023,1)", f64::MAX),
        ("ROUND(0.5,0)", 1.0),
        ("ROUND(2.5,0)", 3.0),
    ]);
}

#[test]
fn digits_past_fifteen_keep_the_fifteen_digit_value() {
    exact(&[
        ("ROUND(0.1+0.2,15)", 0.3),
        ("ROUND(0.1+0.2,16)", 0.3),
        ("ROUND(0.1+0.2,17)", 0.3),
        ("ROUND(1/3,20)", 0.333_333_333_333_333),
        ("ROUND(1+2^-52,15)", 1.0),
    ]);
}
