//! engineering functions, checked against microsoft's documented examples,
//! 30-digit mpmath oracles, and python's integer arithmetic for the base
//! conversions and bitwise functions.

use xlsx_calc::{EvalContext, Value, evaluate_array, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

const RESULT: &str = "ZZ1";

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn eval_in(inputs: &[(&str, CellValue)], formula: &str) -> CellValue {
    let mut sheet = Sheet::new("Sheet1");
    for (address, value) in inputs {
        sheet.set_cell(
            CellRef::parse_a1(address).unwrap(),
            Cell {
                value: value.clone(),
                ..Cell::default()
            },
        );
    }
    sheet.set_cell(
        CellRef::parse_a1(RESULT).unwrap(),
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
        .cell(CellRef::parse_a1(RESULT).unwrap())
        .unwrap()
        .value
        .clone()
}

fn eval(formula: &str) -> CellValue {
    eval_in(&[], formula)
}

/// evaluate in array mode over `inputs`: (rows, cols, values row by row).
fn arrayed_in(inputs: &[(&str, CellValue)], formula: &str) -> (usize, usize, Vec<CellValue>) {
    let mut sheet = Sheet::new("Sheet1");
    for (address, value) in inputs {
        sheet.set_cell(
            CellRef::parse_a1(address).unwrap(),
            Cell {
                value: value.clone(),
                ..Cell::default()
            },
        );
    }
    let book = Workbook {
        sheets: vec![sheet],
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

fn arrayed(formula: &str) -> (usize, usize, Vec<CellValue>) {
    arrayed_in(&[], formula)
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn value(formula: &str) -> f64 {
    match eval(formula) {
        CellValue::Number { value } => value,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

fn exact(formula: &str, want: f64) {
    assert_eq!(eval(formula), number(want), "{formula}");
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

fn says_in(inputs: &[(&str, CellValue)], formula: &str, want: &str) {
    assert_eq!(eval_in(inputs, formula), text(want), "{formula}");
}

fn says(formula: &str, want: &str) {
    says_in(&[], formula, want);
}

fn fails(formula: &str, want: ErrorValue) {
    assert_eq!(eval(formula), CellValue::Error { value: want }, "{formula}");
}

/// the parts of a complex text result, read independently of the engine.
fn parts(formula: &str) -> (f64, f64) {
    let shown = match eval(formula) {
        CellValue::Text { value } => value,
        other => panic!("{formula}: expected text, got {other:?}"),
    };
    let Some(body) = shown.strip_suffix(['i', 'j']) else {
        return (shown.parse().unwrap(), 0.0);
    };
    let bytes = body.as_bytes();
    let split = (1..bytes.len())
        .rev()
        .find(|&k| matches!(bytes[k], b'+' | b'-') && bytes[k - 1] != b'E');
    let (re, im) = match split {
        Some(k) => (body[..k].parse().unwrap(), &body[k..]),
        None => (0.0, body),
    };
    let im = match im {
        "" | "+" => 1.0,
        "-" => -1.0,
        digits => digits.parse().unwrap(),
    };
    (re, im)
}

/// each part within `rel` of the oracle, relative to that part.
fn complex_close(formula: &str, re: f64, im: f64, rel: f64) {
    let (got_re, got_im) = parts(formula);
    for (got, want) in [(got_re, re), (got_im, im)] {
        let scale = want.abs().max(f64::MIN_POSITIVE);
        assert!(
            (got - want).abs() <= rel * scale,
            "{formula}: {got_re}{got_im:+}i vs {re}{im:+}i"
        );
    }
}

#[test]
fn base_conversions_microsoft_examples() {
    exact("BIN2DEC(1100100)", 100.0);
    exact("BIN2DEC(1111111111)", -1.0);
    says("BIN2HEX(11111011,4)", "00FB");
    says("BIN2HEX(1110)", "E");
    says("BIN2HEX(1111111111)", "FFFFFFFFFF");
    says("BIN2OCT(1001,3)", "011");
    says("BIN2OCT(1100100)", "144");
    says("BIN2OCT(1111111111)", "7777777777");
    says("DEC2BIN(9,4)", "1001");
    says("DEC2BIN(-100)", "1110011100");
    says("DEC2HEX(100,4)", "0064");
    says("DEC2HEX(-54)", "FFFFFFFFCA");
    says("DEC2HEX(28)", "1C");
    fails("DEC2HEX(64,1)", ErrorValue::Num);
    says("DEC2OCT(58,3)", "072");
    says("DEC2OCT(-100)", "7777777634");
    says("HEX2BIN(\"F\",8)", "00001111");
    says("HEX2BIN(\"B7\")", "10110111");
    says("HEX2BIN(\"FFFFFFFFFF\")", "1111111111");
    exact("HEX2DEC(\"A5\")", 165.0);
    exact("HEX2DEC(\"FFFFFFFF5B\")", -165.0);
    exact("HEX2DEC(\"3DA408B9\")", 1_034_160_313.0);
    says("HEX2OCT(\"F\",3)", "017");
    says("HEX2OCT(\"3B4E\")", "35516");
    says("HEX2OCT(\"FFFFFFFF00\")", "7777777400");
    says("OCT2BIN(3,3)", "011");
    says("OCT2BIN(7777777000)", "1000000000");
    exact("OCT2DEC(54)", 44.0);
    exact("OCT2DEC(7777777533)", -165.0);
    says("OCT2HEX(100,4)", "0040");
    says("OCT2HEX(7777777533)", "FFFFFFFF5B");
}

#[test]
fn base_conversions_reach_each_range_end() {
    // python int(), format(n & mask, "X"/"o"/"b")
    says("DEC2BIN(511)", "111111111");
    says("DEC2BIN(-512)", "1000000000");
    says("DEC2OCT(536870911)", "3777777777");
    says("DEC2OCT(-536870912)", "4000000000");
    says("DEC2HEX(549755813887)", "7FFFFFFFFF");
    says("DEC2HEX(-549755813888)", "8000000000");
    exact("HEX2DEC(\"8000000000\")", -549_755_813_888.0);
    exact("HEX2DEC(\"7FFFFFFFFF\")", 549_755_813_887.0);
    exact("OCT2DEC(\"4000000000\")", -536_870_912.0);
    exact("BIN2DEC(\"1000000000\")", -512.0);
    says("HEX2BIN(\"FFFFFFFE00\")", "1000000000");
    says("HEX2BIN(\"1FF\")", "111111111");
    says("HEX2OCT(\"FFE0000000\")", "4000000000");
    says("HEX2OCT(\"1FFFFFFF\")", "3777777777");
    says("OCT2BIN(\"777\")", "111111111");
    says("BIN2HEX(\"1000000000\")", "FFFFFFFE00");
    says("OCT2HEX(\"4000000000\")", "FFE0000000");
}

#[test]
fn base_conversions_read_text_numbers_and_either_case() {
    // python int("3da408b9", 16), int("777", 8)
    exact("HEX2DEC(\"3da408b9\")", 1_034_160_313.0);
    exact("OCT2DEC(\"777\")", 511.0);
    exact("BIN2DEC(\"0000000101\")", 5.0);
    says("DEC2HEX(\"255\")", "FF");
    says_in(&[("A1", number(255.0))], "DEC2HEX(A1,4)", "00FF");
    exact("HEX2DEC(\"\")", 0.0);
}

#[test]
fn base_conversions_truncate_number_and_places() {
    says("DEC2BIN(9.9)", "1001");
    says("DEC2BIN(-0.5)", "0");
    says("DEC2BIN(9,4.9)", "1001");
    says("DEC2HEX(255,10)", "00000000FF");
}

#[test]
fn base_conversions_ignore_places_for_negatives() {
    says("DEC2BIN(-100,2)", "1110011100");
    says("BIN2HEX(1111111111,3)", "FFFFFFFFFF");
    says("HEX2OCT(\"FFFFFFFF00\",2)", "7777777400");
}

#[test]
fn base_conversion_errors() {
    fails("BIN2DEC(12)", ErrorValue::Num);
    fails("BIN2DEC(\"11111111111\")", ErrorValue::Num);
    fails("BIN2DEC(-1)", ErrorValue::Num);
    fails("BIN2DEC(TRUE)", ErrorValue::Value);
    fails("HEX2DEC(\"G1\")", ErrorValue::Num);
    fails("HEX2DEC(\" A5\")", ErrorValue::Num);
    fails("OCT2DEC(8)", ErrorValue::Num);
    fails("DEC2BIN(512)", ErrorValue::Num);
    fails("DEC2BIN(-513)", ErrorValue::Num);
    fails("DEC2OCT(536870912)", ErrorValue::Num);
    fails("DEC2HEX(549755813888)", ErrorValue::Num);
    fails("DEC2HEX(-549755813889)", ErrorValue::Num);
    fails("DEC2BIN(\"nine\")", ErrorValue::Value);
    fails("DEC2BIN(9,-1)", ErrorValue::Num);
    fails("DEC2BIN(9,0)", ErrorValue::Num);
    fails("DEC2BIN(9,11)", ErrorValue::Num);
    fails("DEC2BIN(9,\"four\")", ErrorValue::Value);
    fails("DEC2BIN(9,3)", ErrorValue::Num);
    fails("HEX2BIN(\"200\")", ErrorValue::Num);
    fails("HEX2BIN(\"FFFFFFFDFF\")", ErrorValue::Num);
    fails("HEX2OCT(\"20000000\")", ErrorValue::Num);
    fails("HEX2OCT(\"FFDFFFFFFF\")", ErrorValue::Num);
    fails("OCT2BIN(1000)", ErrorValue::Num);
    fails("OCT2BIN(7777776777)", ErrorValue::Num);
    fails("BIN2OCT(1001,1)", ErrorValue::Num);
    fails("DEC2BIN(1/0)", ErrorValue::Div0);
}

#[test]
fn bitwise_microsoft_examples() {
    exact("BITAND(1,5)", 1.0);
    exact("BITAND(13,25)", 9.0);
    exact("BITOR(23,10)", 31.0);
    exact("BITXOR(5,3)", 6.0);
    exact("BITLSHIFT(4,2)", 16.0);
    exact("BITRSHIFT(13,2)", 3.0);
}

#[test]
fn bitwise_matches_integer_arithmetic() {
    // python: (2**48-1) & 123456789, (2**48-1) ^ 1, 1 << 47, (2**48-1) >> 47
    exact("BITAND(281474976710655,123456789)", 123_456_789.0);
    exact("BITXOR(281474976710655,1)", 281_474_976_710_654.0);
    exact("BITOR(140737488355328,1)", 140_737_488_355_329.0);
    exact("BITLSHIFT(1,47)", 140_737_488_355_328.0);
    exact("BITRSHIFT(281474976710655,47)", 1.0);
    exact("BITRSHIFT(5,53)", 0.0);
}

#[test]
fn bit_shifts_run_backwards_for_negative_amounts_and_truncate() {
    exact("BITLSHIFT(13,-2)", 3.0);
    exact("BITRSHIFT(4,-2)", 16.0);
    exact("BITLSHIFT(4,2.9)", 16.0);
    exact("BITRSHIFT(13,-2.9)", 52.0);
}

#[test]
fn bitwise_errors() {
    fails("BITAND(-1,1)", ErrorValue::Num);
    fails("BITOR(1.5,1)", ErrorValue::Num);
    fails("BITXOR(1,281474976710656)", ErrorValue::Num);
    fails("BITAND(\"a\",1)", ErrorValue::Value);
    fails("BITLSHIFT(1,54)", ErrorValue::Num);
    fails("BITRSHIFT(1,-54)", ErrorValue::Num);
    fails("BITLSHIFT(-1,1)", ErrorValue::Num);
    fails("BITRSHIFT(281474976710656,1)", ErrorValue::Num);
    fails("BITLSHIFT(1,\"x\")", ErrorValue::Value);
}

#[test]
fn delta_and_gestep_microsoft_examples() {
    exact("DELTA(5,4)", 0.0);
    exact("DELTA(5,5)", 1.0);
    exact("DELTA(0.5,0)", 0.0);
    exact("DELTA(0)", 1.0);
    exact("GESTEP(5,4)", 1.0);
    exact("GESTEP(5,5)", 1.0);
    exact("GESTEP(-4,-5)", 1.0);
    exact("GESTEP(-1)", 0.0);
    exact("GESTEP(0)", 1.0);
}

#[test]
fn delta_and_gestep_errors() {
    fails("DELTA(\"a\",1)", ErrorValue::Value);
    fails("DELTA(1,\"a\")", ErrorValue::Value);
    fails("GESTEP(\"a\")", ErrorValue::Value);
    fails("GESTEP(1,\"a\")", ErrorValue::Value);
    fails("DELTA(NA())", ErrorValue::NA);
}

#[test]
fn complex_microsoft_examples() {
    says("COMPLEX(3,4)", "3+4i");
    says("COMPLEX(3,4,\"j\")", "3+4j");
    says("COMPLEX(0,1)", "i");
    says("COMPLEX(1,0)", "1");
}

#[test]
fn complex_text_drops_zero_parts_and_unit_coefficients() {
    says("COMPLEX(0,0)", "0");
    says("COMPLEX(0,-1)", "-i");
    says("COMPLEX(-1.5,1,\"j\")", "-1.5+j");
    says("COMPLEX(2,-1)", "2-i");
    says("COMPLEX(0,-2.5)", "-2.5i");
    says("COMPLEX(3,4,\"\")", "3+4i");
}

#[test]
fn complex_text_keeps_fifteen_significant_digits() {
    says("COMPLEX(1/3,0)", "0.333333333333333");
    says("COMPLEX(2/3,-2/3)", "0.666666666666667-0.666666666666667i");
    says("COMPLEX(123456789012345,1)", "123456789012345+i");
    says("COMPLEX(1E+15,0)", "1E+15");
    says("COMPLEX(1234567890123456,0)", "1.23456789012346E+15");
    says("COMPLEX(0.0001,0)", "0.0001");
    says("COMPLEX(0.000012345,0)", "1.2345E-05");
    says("COMPLEX(1E-300,2.5E+100)", "1E-300+2.5E+100i");
    says("COMPLEX(-12.5,0.125)", "-12.5+0.125i");
}

#[test]
fn complex_errors() {
    fails("COMPLEX(\"a\",1)", ErrorValue::Value);
    fails("COMPLEX(1,\"b\")", ErrorValue::Value);
    fails("COMPLEX(1,1,\"k\")", ErrorValue::Value);
    fails("COMPLEX(1,1,\"I\")", ErrorValue::Value);
    fails("COMPLEX(1,1,\"J\")", ErrorValue::Value);
    fails("COMPLEX(TRUE,1)", ErrorValue::Value);
}

#[test]
fn complex_parts_microsoft_examples() {
    exact("IMABS(\"5+12i\")", 13.0);
    exact("IMAGINARY(\"3+4i\")", 4.0);
    exact("IMAGINARY(\"0-j\")", -1.0);
    exact("IMAGINARY(4)", 0.0);
    exact("IMREAL(\"6-9i\")", 6.0);
    near("IMARGUMENT(\"3+4i\")", 0.92729522, 5e-9);
    says("IMCONJUGATE(\"3+4i\")", "3-4i");
}

#[test]
fn complex_parts_match_mpmath() {
    // mpmath 30 digits
    close("IMABS(\"1.5-2.25i\")", 2.704_163_456_597_992, 1e-15);
    close("IMARGUMENT(\"1.5-2.25i\")", -0.982_793_723_247_329, 1e-15);
    close("IMABS(\"-0.75+0.4j\")", 0.85, 1e-15);
    close("IMARGUMENT(\"-0.75+0.4j\")", 2.651_635_327_336_065, 1e-15);
}

#[test]
fn complex_text_is_read_the_way_excel_writes_it() {
    exact("IMREAL(\"1E2+3i\")", 100.0);
    exact("IMAGINARY(\"1E2+3i\")", 3.0);
    exact("IMAGINARY(\"1+2E-3i\")", 0.002);
    exact("IMAGINARY(\"1E+2i\")", 100.0);
    exact("IMAGINARY(\"4j\")", 4.0);
    exact("IMAGINARY(\"-i\")", -1.0);
    exact("IMAGINARY(\"+i\")", 1.0);
    exact("IMAGINARY(\"i\")", 1.0);
    exact("IMREAL(\"-2.5e-1-3i\")", -0.25);
    exact("IMREAL(\".5+i\")", 0.5);
    exact("IMREAL(\"7\")", 7.0);
    exact("IMAGINARY(\"7\")", 0.0);
    exact("IMREAL(\"-i\")", 0.0);
    says("IMCONJUGATE(\"3+4j\")", "3-4j");
    says("IMCONJUGATE(5)", "5");
}

#[test]
fn unreadable_complex_text_is_num() {
    fails("IMREAL(\"3 + 4i\")", ErrorValue::Num);
    fails("IMREAL(\"3+4I\")", ErrorValue::Num);
    fails("IMREAL(\"3+4ij\")", ErrorValue::Num);
    fails("IMREAL(\"3+-4i\")", ErrorValue::Num);
    fails("IMREAL(\"4+3\")", ErrorValue::Num);
    fails("IMREAL(\"abc\")", ErrorValue::Num);
    fails("IMREAL(\"i3\")", ErrorValue::Num);
    fails("IMREAL(\"1E\")", ErrorValue::Num);
    fails("IMABS(\"inf\")", ErrorValue::Num);
    fails("IMREAL(TRUE)", ErrorValue::Value);
    fails("IMREAL(NA())", ErrorValue::NA);
}

#[test]
fn mixed_units_are_value_errors() {
    fails("IMSUM(\"1+i\",\"1+j\")", ErrorValue::Value);
    fails("IMSUB(\"1+i\",\"j\")", ErrorValue::Value);
    fails("IMDIV(\"1+j\",\"2i\")", ErrorValue::Value);
    fails("IMPRODUCT(\"1+i\",2,\"3j\")", ErrorValue::Value);
    says("IMSUM(\"1+j\",2)", "3+j");
    says("IMSUB(\"5+3j\",2)", "3+3j");
}

#[test]
fn complex_arithmetic_microsoft_examples() {
    says("IMDIV(\"-238+240i\",\"10+24i\")", "5+12i");
    says("IMPRODUCT(\"3+4i\",\"5-3i\")", "27+11i");
    says("IMPRODUCT(\"1+2i\",30)", "30+60i");
    says("IMSUB(\"13+4i\",\"5+3i\")", "8+i");
    says("IMSUM(\"3+4i\",\"5-3i\")", "8+i");
}

#[test]
fn complex_arithmetic_matches_mpmath() {
    // mpmath 30 digits
    complex_close(
        "IMDIV(\"1.5-2.25i\",\"-0.75+0.4i\")",
        -2.802_768_166_089_965_4,
        1.505_190_311_418_685,
        1e-13,
    );
    complex_close(
        "IMPRODUCT(\"1.5-2.25i\",\"-0.75+0.4i\",\"3+i\")",
        -2.9625,
        6.6375,
        1e-13,
    );
}

#[test]
fn imsum_and_improduct_take_ranges_and_arrays() {
    let cells = [
        ("A1", text("1+i")),
        ("A2", number(2.0)),
        ("A4", text("-3i")),
    ];
    says_in(&cells, "IMSUM(A1:A4)", "3-2i");
    says_in(&cells, "IMPRODUCT(A1:A4)", "6-6i");
    says_in(&cells, "IMSUM(A1:A2,\"1\",{\"i\",2})", "6+2i");
    says("IMSUM({\"1+i\",\"2-3i\"})", "3-2i");
    says("IMPRODUCT({\"i\";\"i\"})", "-1");
}

#[test]
fn complex_arithmetic_errors() {
    fails("IMDIV(\"1+i\",0)", ErrorValue::Num);
    fails("IMDIV(\"1+i\",\"x\")", ErrorValue::Num);
    fails("IMSUB(\"x\",1)", ErrorValue::Num);
    fails("IMSUM(\"1+i\",TRUE)", ErrorValue::Value);
    fails("IMPOWER(\"2+3i\",\"x\")", ErrorValue::Value);
    fails("IMPOWER(0,0)", ErrorValue::Num);
    fails("IMPOWER(0,-1)", ErrorValue::Num);
    says("IMPOWER(0,2)", "0");
}

#[test]
fn complex_functions_microsoft_examples() {
    says("IMPOWER(\"2+3i\",3)", "-46+9.00000000000001i");
    says("IMSQRT(\"1+i\")", "1.09868411346781+0.455089860562227i");
    says("IMEXP(\"1+i\")", "1.46869393991589+2.28735528717884i");
    says("IMLN(\"3+4i\")", "1.6094379124341+0.927295218001612i");
    says("IMLOG10(\"3+4i\")", "0.698970004336019+0.402719196273373i");
    says("IMLOG2(\"3+4i\")", "2.32192809488736+1.33780421245098i");
    says("IMSIN(\"4+3i\")", "-7.61923172032141-6.548120040911i");
    says("IMCOS(\"1+i\")", "0.833730025131149-0.988897705762865i");
    says("IMSINH(\"4+3i\")", "-27.0168132580039+3.85373803791938i");
    says("IMCOSH(\"4+3i\")", "-27.0349456030742+3.85115333481178i");
    says("IMTAN(\"4+3i\")", "0.00490825806749606+1.00070953606723i");
    says("IMCSC(\"4+3i\")", "-0.0754898329158637+0.0648774713706355i");
    says(
        "IMSECH(\"4+3i\")",
        "-0.0362534969158689-0.00516434460775318i",
    );
    says("IMCSCH(\"4+3i\")", "-0.036275889628626-0.0051744731840194i");
}

/// excel's double-angle secant lands one unit above the correctly rounded
/// -0.065294027857947 in the fifteenth digit.
#[test]
fn imsec_keeps_excels_fifteenth_digit() {
    says("IMSEC(\"4+3i\")", "-0.0652940278579471-0.0752249603027732i");
}

#[test]
fn imcot_matches_mpmath() {
    // mpmath 30 digits
    complex_close(
        "IMCOT(\"4+3i\")",
        0.004_901_182_394_304_473,
        -0.999_266_927_805_901_5,
        1e-13,
    );
}

/// the polar forms keep the trigonometric residue where an exact answer is
/// real or imaginary.
#[test]
fn polar_forms_keep_the_residue_excel_prints() {
    says("IMSQRT(\"-4\")", "1.22464679914735E-16+2i");
    says("IMPOWER(\"i\",2)", "-1+1.22464679914735E-16i");
    says("IMSQRT(0)", "0");
    says("IMSQRT(\"4\")", "2");
}

#[test]
fn complex_functions_match_mpmath() {
    // mpmath 30 digits, at 1.5-2.25i and -0.75+0.4i
    let cases: &[(&str, [f64; 4])] = &[
        (
            "IMSIN",
            [
                4.784_552_064_541_835,
                -0.331_840_118_511_466_4,
                -0.736_900_831_035_449_2,
                0.300_542_904_653_953_2,
            ],
        ),
        (
            "IMCOS",
            [
                0.339_295_764_714_918_5,
                4.679_416_866_449_42,
                0.791_008_620_921_217_7,
                0.279_984_706_036_931_75,
            ],
        ),
        (
            "IMTAN",
            [
                0.003_205_515_147_866_466,
                -1.022_235_159_497_119_5,
                -0.708_361_536_648_927_3,
                0.630_679_980_048_104_7,
            ],
        ),
        (
            "IMCOT",
            [
                0.003_067_552_006_674_880_4,
                0.978_238_869_623_221_5,
                -0.787_476_721_125_604_3,
                -0.701_119_099_601_799_6,
            ],
        ),
        (
            "IMSEC",
            [
                0.015_414_082_359_451_234,
                -0.212_584_194_896_332_98,
                1.123_454_432_840_605_4,
                -0.397_656_929_147_543_5,
            ],
        ),
        (
            "IMCSC",
            [
                0.208_005_406_014_424_17,
                0.014_426_541_429_947_205,
                -1.163_499_081_221_687_7,
                -0.474_529_785_698_871_3,
            ],
        ),
        (
            "IMSINH",
            [
                -1.337_557_189_096_011_4,
                -1.830_346_869_722_198_2,
                -0.757_403_866_502_019_6,
                0.504_173_418_533_575_5,
            ],
        ),
        (
            "IMCOSH",
            [
                -1.477_721_670_135_155_5,
                -1.656_735_272_693_395_6,
                1.192_482_273_103_374_8,
                -0.320_225_218_603_117_84,
            ],
        ),
        (
            "IMSECH",
            [
                -0.299_836_000_753_019_86,
                0.336_158_620_740_411_7,
                0.782_182_210_292_394,
                0.210_044_606_052_301_82,
            ],
        ),
        (
            "IMCSCH",
            [
                -0.260_264_178_877_016_6,
                0.356_152_042_687_851_8,
                -0.914_903_107_078_990_7,
                -0.609_014_354_855_785_1,
            ],
        ),
        (
            "IMEXP",
            [
                -2.815_278_859_231_167,
                -3.487_082_142_415_594,
                0.435_078_406_601_355_2,
                0.183_948_199_930_457_7,
            ],
        ),
        (
            "IMLN",
            [
                0.994_792_606_278_987_4,
                -0.982_793_723_247_329,
                -0.162_518_929_497_774_9,
                2.651_635_327_336_065,
            ],
        ),
        (
            "IMLOG10",
            [
                0.432_032_939_545_118_43,
                -0.426_821_890_855_466_6,
                -0.070_581_074_285_707_26,
                1.151_590_590_681_775_9,
            ],
        ),
        (
            "IMLOG2",
            [
                1.435_182_359_791_702_3,
                -1.417_871_630_745_722,
                -0.234_465_253_637_022_92,
                3.825_501_136_993_724,
            ],
        ),
        (
            "IMSQRT",
            [
                1.449_855_761_204_884_8,
                -0.775_939_255_547_105_3,
                0.223_606_797_749_979,
                0.894_427_190_999_915_9,
            ],
        ),
    ];
    for (name, [re1, im1, re2, im2]) in cases {
        complex_close(&format!("{name}(\"1.5-2.25i\")"), *re1, *im1, 1e-13);
        complex_close(&format!("{name}(\"-0.75+0.4i\")"), *re2, *im2, 1e-13);
    }
    complex_close(
        "IMPOWER(\"1.5-2.25i\",2.5)",
        -9.315_309_303_331_7,
        -7.604_197_231_906_739,
        1e-13,
    );
    complex_close(
        "IMPOWER(\"1.5-2.25i\",-3)",
        -0.049_629_966_789_730_1,
        0.009_710_210_893_642_846,
        1e-13,
    );
    complex_close(
        "IMPOWER(\"-0.75+0.4i\",2.5)",
        0.626_658_050_694_316_1,
        0.225_842_865_727_478_17,
        1e-13,
    );
    complex_close(
        "IMPOWER(\"-0.75+0.4i\",-3)",
        -0.164_059_603_516_824_63,
        -1.620_047_155_535_837,
        1e-13,
    );
}

#[test]
fn complex_function_errors() {
    fails("IMLN(0)", ErrorValue::Num);
    fails("IMLOG10(0)", ErrorValue::Num);
    fails("IMLOG2(0)", ErrorValue::Num);
    fails("IMARGUMENT(0)", ErrorValue::Div0);
    fails("IMCOT(0)", ErrorValue::Num);
    fails("IMCSC(0)", ErrorValue::Num);
    fails("IMCSCH(0)", ErrorValue::Num);
    fails("IMEXP(1000)", ErrorValue::Num);
    fails("IMSIN(\"x\")", ErrorValue::Num);
    fails("IMTAN(TRUE)", ErrorValue::Value);
    fails("IMSECH(FALSE)", ErrorValue::Value);
}

#[test]
fn bessel_microsoft_examples() {
    near("BESSELI(1.5,1)", 0.981_666_428, 1e-9);
    near("BESSELY(2.5,1)", 0.145_918_138, 1e-9);
}

/// microsoft prints BESSELJ(1.9,2) as 0.329925829 and BESSELK(1.5,1) as
/// 0.277387804, about 1e-7 and 4e-9 from the true 0.3299257277 and
/// 0.2773878005; these match the true values (mpmath 30 digits) and so sit
/// that far from the printed ones.
#[test]
fn bessel_j_and_k_match_the_true_value_of_microsofts_examples() {
    close("BESSELJ(1.9,2)", 0.329_925_727_692_387_2, 1e-12);
    close("BESSELK(1.5,1)", 0.277_387_800_456_843_8, 1e-12);
    near("BESSELJ(1.9,2)", 0.329_925_829, 1.1e-7);
    near("BESSELK(1.5,1)", 0.277_387_804, 4e-9);
}

#[test]
fn besselj_matches_mpmath() {
    // mpmath 30 digits
    close("BESSELJ(0.5,0)", 0.938_469_807_240_812_9, 1e-12);
    close("BESSELJ(10,3)", 0.058_379_379_305_186_81, 1e-12);
    close("BESSELJ(30,2)", 0.078_451_246_073_265_35, 1e-12);
    close("BESSELJ(-7.3,3)", 0.228_101_889_059_524_63, 1e-12);
    close("BESSELJ(1,20)", 3.873_503_008_524_658e-25, 1e-12);
    close("BESSELJ(100.5,1)", -0.057_791_123_996_932_02, 1e-12);
    close("BESSELJ(1234.5,5)", 0.018_084_932_113_758_09, 1e-11);
    exact("BESSELJ(0,0)", 1.0);
    exact("BESSELJ(0,3)", 0.0);
}

#[test]
fn bessely_matches_mpmath() {
    // mpmath 30 digits
    close("BESSELY(0.1,0)", -1.534_238_651_350_366_8, 1e-12);
    close("BESSELY(10,5)", 0.135_403_047_689_362_3, 1e-12);
    close("BESSELY(30,1)", 0.084_425_570_661_747_23, 1e-12);
    close("BESSELY(50.5,3)", 0.100_480_953_225_805_83, 1e-12);
    close("BESSELY(0.001,2)", -1_273_239.863_045_667_4, 1e-12);
    close("BESSELY(2,10)", -129_184.542_208_039_28, 1e-12);
}

#[test]
fn besseli_matches_mpmath() {
    // mpmath 30 digits
    close("BESSELI(0.5,0)", 1.063_483_370_741_323_5, 1e-12);
    close("BESSELI(10,2)", 2_281.518_967_726_003_5, 1e-12);
    close("BESSELI(-3,1)", -3.953_370_217_402_609_4, 1e-12);
    close("BESSELI(100,0)", 1.073_751_707_131_073_8e42, 1e-12);
    close("BESSELI(5,30)", 3.997_844_971_250_565e-21, 1e-12);
    exact("BESSELI(0,0)", 1.0);
}

#[test]
fn besselk_matches_mpmath() {
    // mpmath 30 digits
    close("BESSELK(0.01,0)", 4.721_244_730_161_095, 1e-12);
    close("BESSELK(1,2)", 1.624_838_898_635_177_5, 1e-12);
    close("BESSELK(10,1)", 1.864_877_345_382_558_5e-5, 1e-12);
    close("BESSELK(30,0)", 2.132_477_496_463_056_4e-14, 1e-12);
    close("BESSELK(0.5,5)", 12_097.979_476_096_393, 1e-12);
    close("BESSELK(200,3)", 1.253_501_761_543_211e-88, 1e-12);
}

#[test]
fn bessel_orders_are_truncated() {
    for formula in [
        "BESSELJ(1.9,2)",
        "BESSELY(2.5,1)",
        "BESSELI(1.5,1)",
        "BESSELK(1.5,1)",
    ] {
        assert!(
            matches!(eval(formula), CellValue::Number { .. }),
            "{formula}"
        );
    }
    assert_eq!(eval("BESSELJ(1.9,2.9)"), eval("BESSELJ(1.9,2)"));
    assert_eq!(eval("BESSELY(2.5,1.5)"), eval("BESSELY(2.5,1)"));
    assert_eq!(eval("BESSELI(1.5,1.99)"), eval("BESSELI(1.5,1)"));
    assert_eq!(eval("BESSELK(1.5,1.2)"), eval("BESSELK(1.5,1)"));
}

#[test]
fn bessel_errors() {
    fails("BESSELJ(1,-1)", ErrorValue::Num);
    fails("BESSELI(1,-1)", ErrorValue::Num);
    fails("BESSELK(1,-1)", ErrorValue::Num);
    fails("BESSELY(1,-1)", ErrorValue::Num);
    fails("BESSELK(0,1)", ErrorValue::Num);
    fails("BESSELK(-1,1)", ErrorValue::Num);
    fails("BESSELY(0,1)", ErrorValue::Num);
    fails("BESSELY(-2.5,1)", ErrorValue::Num);
    fails("BESSELJ(\"a\",1)", ErrorValue::Value);
    fails("BESSELY(1,\"a\")", ErrorValue::Value);
    fails("BESSELI(TRUE,1)", ErrorValue::Value);
    fails("BESSELY(1E-300,200)", ErrorValue::Num);
}

#[test]
fn scalar_engineering_functions_apply_per_element() {
    let want = vec![number(1.0), number(2.0), number(5.0)];
    assert_eq!(arrayed("BIN2DEC({\"1\",\"10\",\"101\"})"), (1, 3, want));
    let want = vec![text("1"), text("A"), text("FF")];
    assert_eq!(arrayed("DEC2HEX({1;10;255})"), (3, 1, want));
    let want = vec![number(5.0), number(13.0)];
    assert_eq!(arrayed("IMABS({\"3+4i\",\"5+12i\"})"), (1, 2, want));
    let want = vec![text("3-4i"), text("-2j")];
    assert_eq!(arrayed("IMCONJUGATE({\"3+4i\",\"2j\"})"), (1, 2, want));
    let want = vec![text("-46+9.00000000000001i"), text("4")];
    assert_eq!(arrayed("IMPOWER({\"2+3i\",\"4\"},{3,1})"), (1, 2, want));
}

#[test]
fn scalar_engineering_functions_lift_over_ranges() {
    let cells = [
        ("A1", text("3+4i")),
        ("A2", text("5+12i")),
        ("A3", number(2.0)),
        ("B1", number(1.0)),
        ("B2", number(2.0)),
    ];
    let want = vec![number(5.0), number(13.0), number(2.0)];
    assert_eq!(arrayed_in(&cells, "IMABS(A1:A3)"), (3, 1, want));
    let want = vec![text("0001"), text("0010")];
    assert_eq!(arrayed_in(&cells, "DEC2BIN(B1:B2,4)"), (2, 1, want));
    let want = vec![eval("BESSELK(1,2)"), eval("BESSELK(2,2)")];
    assert_eq!(arrayed_in(&cells, "BESSELK(B1:B2,2)"), (2, 1, want));
}

#[test]
fn imsum_and_improduct_take_arrays_whole() {
    let cells = [("A1", text("3+4i")), ("A2", text("5+12i"))];
    let want = vec![text("8+16i")];
    assert_eq!(arrayed_in(&cells, "IMSUM(A1:A2)"), (1, 1, want));
    let want = vec![text("-33+56i")];
    assert_eq!(arrayed_in(&cells, "IMPRODUCT(A1:A2)"), (1, 1, want));
    let want = vec![text("3+i")];
    assert_eq!(arrayed("IMSUM({\"1+i\",\"2\"})"), (1, 1, want));
    let want = vec![text("-1")];
    assert_eq!(arrayed("IMPRODUCT({\"i\";\"i\"})"), (1, 1, want));
}

#[test]
fn scalar_arguments_broadcast_against_arrays() {
    let want = vec![text("0001"), text("0010"), text("0011")];
    assert_eq!(arrayed("DEC2BIN({1,2,3},4)"), (1, 3, want));
    let want = vec![number(1.0), number(1.0), number(1.0), number(3.0)];
    assert_eq!(arrayed("BITAND({1,3},{5;7})"), (2, 2, want));
    let want = vec![number(0.0), number(1.0), number(0.0)];
    assert_eq!(arrayed("DELTA({1,2,3},2)"), (1, 3, want));
    let want = vec![text("1"), text("i"), text("1+i")];
    assert_eq!(arrayed("COMPLEX({1,0,1},{0,1,1})"), (1, 3, want));
    let want = vec![text("3+4j"), text("3-4j")];
    assert_eq!(arrayed("COMPLEX(3,{4,-4},\"j\")"), (1, 2, want));
    let want = vec![text("8+i"), text("6+3i")];
    assert_eq!(arrayed("IMSUB({\"9+4i\",\"7+6i\"},\"1+3i\")"), (1, 2, want));
    let want = vec![eval("BESSELJ(1.9,2)"), eval("BESSELJ(2.5,2)")];
    assert_eq!(arrayed("BESSELJ({1.9,2.5},2)"), (1, 2, want));
}

#[test]
fn errors_stay_in_their_own_elements() {
    let want = vec![
        number(1.0),
        error(ErrorValue::Num),
        error(ErrorValue::Value),
    ];
    assert_eq!(arrayed("BIN2DEC({\"1\",\"2\",TRUE})"), (1, 3, want));
    let want = vec![number(5.0), error(ErrorValue::Num)];
    assert_eq!(arrayed("IMABS({\"3+4i\",\"x\"})"), (1, 2, want));
    let want = vec![text("0001"), error(ErrorValue::Num)];
    assert_eq!(arrayed("DEC2BIN({1,600},4)"), (1, 2, want));
    let want = vec![eval("BESSELJ(1.9,2)"), error(ErrorValue::Num)];
    assert_eq!(arrayed("BESSELJ(1.9,{2,-1})"), (1, 2, want));
    let want = vec![text("1+i"), error(ErrorValue::Num)];
    assert_eq!(arrayed("IMDIV(\"1+i\",{1,0})"), (1, 2, want));
    let want = vec![text("1"), error(ErrorValue::Value)];
    assert_eq!(arrayed("IMSUB({\"1+i\",\"1+j\"},\"i\")"), (1, 2, want));
    let want = vec![number(2.0), error(ErrorValue::NA)];
    assert_eq!(arrayed("IMREAL(IF({1,0},\"2+i\",NA()))"), (1, 2, want));
}

#[test]
fn negative_conversions_still_check_places() {
    fails("DEC2BIN(-1,0)", ErrorValue::Num);
    fails("DEC2BIN(-1,-1)", ErrorValue::Num);
    fails("DEC2BIN(-1,11)", ErrorValue::Num);
    fails("DEC2HEX(-54,0)", ErrorValue::Num);
    fails("BIN2HEX(1111111111,0)", ErrorValue::Num);
    fails("HEX2OCT(\"FFFFFFFF00\",11)", ErrorValue::Num);
    says("DEC2BIN(-1,10)", "1111111111");
    says("DEC2OCT(-100,1)", "7777777634");
}

#[test]
fn bessel_orders_past_32_bits_are_num() {
    fails("BESSELJ(1.5E19,2E19)", ErrorValue::Num);
    fails(
        "BESSELJ(18446744073709551616,18446744073709551616)",
        ErrorValue::Num,
    );
    fails("BESSELI(1,2147483648)", ErrorValue::Num);
    fails("BESSELK(1,2147483648)", ErrorValue::Num);
    fails("BESSELY(1,2147483648)", ErrorValue::Num);
    exact("BESSELJ(1,2147483647)", 0.0);
    exact("BESSELI(1,2147483647)", 0.0);
    fails("BESSELY(1,2147483647)", ErrorValue::Num);
}

#[test]
fn besseli_keeps_large_orders_and_arguments_in_range() {
    // mpmath 30 digits
    close("BESSELI(6000,9000)", 1.928_769_140_533_089_3e25, 1e-12);
    close("BESSELI(4600,6900)", 7.573_779_520_668_245e18, 1e-12);
    exact("BESSELI(1000,150000)", 0.0);
}

#[test]
fn besselk_reaches_the_smallest_arguments() {
    // mpmath 30 digits
    close("BESSELK(1E-307,1)", 1e307, 1e-12);
    close("BESSELK(1E-308,0)", 709.312_140_157_824_5, 1e-12);
    close("BESSELK(1E-10,1)", 1e10, 1e-12);
}

#[test]
fn imdiv_scales_huge_and_tiny_divisors() {
    says("IMDIV(\"1+i\",\"1E200+1E200i\")", "1E-200");
    says("IMDIV(\"1E-200\",\"1E-200\")", "1");
    says("IMDIV(\"3+4i\",\"1E-300+2E-300i\")", "2.2E+300-4E+299i");
    says("IMDIV(\"1E200+1E200i\",\"1E200-1E200i\")", "i");
}

#[test]
fn reciprocal_trigonometry_keeps_small_denominators() {
    // mpmath 30 digits
    says("IMCOT(1E-9)", "1000000000");
    says("IMCSC(1E-9)", "1000000000");
    says("IMCSCH(1E-9)", "1000000000");
    says("IMCSC(\"1E-9+1E-9i\")", "500000000-500000000i");
    says("IMCOT(\"1E-9i\")", "-1000000000i");
}

#[test]
fn reciprocal_trigonometry_survives_large_arguments() {
    // mpmath 30 digits: sech(400) = csch(400) = 3.83033919342801E-174
    says("IMTAN(\"400i\")", "i");
    says("IMCOT(\"400i\")", "-i");
    says("IMTAN(\"1-400i\")", "-i");
    says("IMSEC(\"400i\")", "3.83033919342801E-174");
    says("IMCSC(\"400i\")", "-3.83033919342801E-174i");
    says("IMSECH(400)", "3.83033919342801E-174");
    says("IMCSCH(400)", "3.83033919342801E-174");
    says("IMSECH(1000)", "0");
}

#[test]
fn complex_logarithms_and_roots_take_the_largest_operands() {
    // mpmath 30 digits
    let z = "\"1.5E+308+1.5E+308i\"";
    let quarter = std::f64::consts::FRAC_PI_4;
    let log = 709.948_247_340_554_2;
    complex_close(&format!("IMLN({z})"), log, quarter, 1e-13);
    let (re, im) = (1.345_607_733_249_115e154, 5.573_689_727_459_013_4e153);
    complex_close(&format!("IMSQRT({z})"), re, im, 1e-13);
    complex_close(&format!("IMPOWER({z},0.5)"), re, im, 1e-13);
    let (re, im) = (308.326_606_256_887_66, 0.341_094_088_460_460_35);
    complex_close(&format!("IMLOG10({z})"), re, im, 1e-13);
    let (re, im) = (1_024.238_815_726_028_9, 1.133_090_035_456_798_5);
    complex_close(&format!("IMLOG2({z})"), re, im, 1e-13);
    fails(&format!("IMABS({z})"), ErrorValue::Num);
}

#[test]
fn besselk_keeps_its_seeds_scaled_past_exp_underflow() {
    // mpmath 30 digits
    close("BESSELK(1000,1000)", 1.298_180_251_466_701e-233, 1e-12);
    close("BESSELK(700,3)", 4.699_871_529_146_931e-306, 1e-12);
    exact("BESSELK(1000,0)", 0.0);
    exact("BESSELK(1E200,0)", 0.0);
    exact("BESSELK(1E308,5)", 0.0);
}

#[test]
fn negative_fractional_bessel_orders_are_num() {
    fails("BESSELJ(1,-0.5)", ErrorValue::Num);
    fails("BESSELI(1,-0.5)", ErrorValue::Num);
    fails("BESSELK(1,-0.5)", ErrorValue::Num);
    fails("BESSELY(1,-0.5)", ErrorValue::Num);
}

#[test]
fn engineering_functions_at_extreme_operands() {
    // mpmath 30 digits
    says("DEC2HEX(549755813887.75)", "7FFFFFFFFF");
    fails("DEC2BIN(1E200)", ErrorValue::Num);
    fails("DEC2HEX(1E15+0.75)", ErrorValue::Num);
    fails("HEX2DEC(1E15)", ErrorValue::Num);
    fails("BIN2DEC(1E-200)", ErrorValue::Num);
    fails("BITAND(9007199254740992,1)", ErrorValue::Num);
    fails("BITLSHIFT(1,1E200)", ErrorValue::Num);
    exact("BITLSHIFT(281474976710655,53)", 2.535_301_200_456_449_8e30);
    says("IMPRODUCT(\"1E200+1E200i\",\"1E-200-1E-200i\")", "2");
    fails("IMSUM(\"1E308\",\"1E308\")", ErrorValue::Num);
    fails("IMPRODUCT(\"1E200i\",\"1E200i\")", ErrorValue::Num);
    says("IMEXP(-1000)", "0");
    fails("IMEXP(710)", ErrorValue::Num);
    says("COMPLEX(1E308,-1E308)", "1E+308-1E+308i");
    says("COMPLEX(-0,1E-308)", "1E-308i");
    close("BESSELY(1E-308,0)", -451.562_133_204_836_2, 1e-12);
    close("BESSELY(1E-308,1)", -6.366_197_723_675_814e307, 1e-12);
    exact("BESSELJ(1E-308,0)", 1.0);
    exact("BESSELI(1E-308,0)", 1.0);
    fails("BESSELI(1E200,0)", ErrorValue::Num);
}
