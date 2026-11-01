//! math and trig functions past the core set, checked against microsoft's
//! documented examples, 30-digit mpmath oracles (computed from the exact
//! double each input parses to) and exact integer arithmetic.

use xlsx_calc::{EvalContext, Value, evaluate_array, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{Cell, CellRef, CellValue, DateSystem, ErrorValue, Sheet, SheetId, Workbook};

const RESULT: &str = "ZZ1";

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn cells(inputs: &[(&str, CellValue)]) -> Vec<(String, CellValue)> {
    inputs
        .iter()
        .map(|(address, value)| (address.to_string(), value.clone()))
        .collect()
}

fn book(inputs: &[(String, CellValue)], date_system: DateSystem) -> Workbook {
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
    Workbook {
        sheets: vec![sheet],
        date_system,
        ..Workbook::default()
    }
}

fn eval_with(inputs: &[(String, CellValue)], date_system: DateSystem, formula: &str) -> CellValue {
    let mut book = book(inputs, date_system);
    book.sheets[0].set_cell(
        CellRef::parse_a1(RESULT).unwrap(),
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
    rebuild_and_recalc_all(&mut book, None);
    book.sheets[0]
        .cell(CellRef::parse_a1(RESULT).unwrap())
        .unwrap()
        .value
        .clone()
}

fn eval_in(inputs: &[(&str, CellValue)], formula: &str) -> CellValue {
    eval_with(&cells(inputs), DateSystem::V1900, formula)
}

fn eval(formula: &str) -> CellValue {
    eval_in(&[], formula)
}

fn value_of(result: CellValue, formula: &str) -> f64 {
    match result {
        CellValue::Number { value } => value,
        other => panic!("{formula} gave {other:?}"),
    }
}

fn assert_close_in(inputs: &[(&str, CellValue)], formula: &str, expected: f64, relative: f64) {
    let actual = value_of(eval_in(inputs, formula), formula);
    let scale = expected.abs().max(f64::MIN_POSITIVE);
    assert!(
        (actual - expected).abs() <= relative * scale,
        "{formula}: {actual:e} vs {expected:e}"
    );
}

/// within 1e-12 relative.
fn assert_close(formula: &str, expected: f64) {
    assert_close_in(&[], formula, expected, 1e-12);
}

/// a published result, which microsoft prints rounded to `places` decimals.
fn assert_published(formula: &str, expected: f64, places: i32) {
    let actual = value_of(eval(formula), formula);
    assert!(
        (actual - expected).abs() <= 0.5 * 10f64.powi(-places),
        "{formula}: {actual} vs published {expected}"
    );
}

fn assert_exact(formula: &str, expected: f64) {
    assert_eq!(eval(formula), number(expected), "{formula}");
}

fn assert_text(formula: &str, expected: &str) {
    assert_eq!(eval(formula), text(expected), "{formula}");
}

fn assert_error(formula: &str, expected: ErrorValue) {
    assert_eq!(eval(formula), error(expected), "{formula}");
}

/// evaluate in array mode, as (rows, cols, numbers).
fn block(formula: &str, seed: Option<u64>) -> (usize, usize, Vec<f64>) {
    let book = book(&[], DateSystem::V1900);
    let expression = parse_formula(formula).unwrap();
    let mut context = EvalContext::new(&book, SheetId(0));
    context.rand_seed = seed;
    let unwrap = |value: CellValue| match value {
        CellValue::Number { value } => value,
        other => panic!("{formula} held {other:?}"),
    };
    match evaluate_array(&expression, &context) {
        Value::Array(array) => {
            let mut values = Vec::new();
            for row in 0..array.rows() {
                for col in 0..array.cols() {
                    values.push(unwrap(array.at(row, col)));
                }
            }
            (array.rows(), array.cols(), values)
        }
        other => (1, 1, vec![unwrap(other.into_scalar())]),
    }
}

fn block_error(formula: &str) -> CellValue {
    let book = book(&[], DateSystem::V1900);
    let expression = parse_formula(formula).unwrap();
    let context = EvalContext::new(&book, SheetId(0));
    evaluate_array(&expression, &context).into_scalar()
}

fn assert_block(formula: &str, rows: usize, cols: usize, expected: &[f64], relative: f64) {
    let (r, c, values) = block(formula, None);
    assert_eq!((r, c), (rows, cols), "{formula}");
    for (actual, expected) in values.iter().zip(expected) {
        assert!(
            (actual - expected).abs() <= relative * expected.abs(),
            "{formula}: {actual:e} vs {expected:e}"
        );
    }
}

#[test]
fn inverse_hyperbolic_functions_match_documentation_and_oracles() {
    assert_exact("ACOSH(1)", 0.0);
    assert_published("ACOSH(10)", 2.9932228, 7);
    assert_published("ASINH(-2.5)", -1.647231146, 9);
    assert_published("ASINH(10)", 2.99822295, 8);
    assert_published("ATANH(0.76159416)", 1.00000001, 8);
    assert_published("ATANH(-0.1)", -0.100335348, 9);
    // mpmath 30 digits
    assert_close("ACOSH(1.5)", 0.9624236501192069);
    assert_close("ACOSH(1.0000001)", 0.0004472135919037347);
    assert_close("ACOSH(1E10)", 23.7189981105004);
    assert_close("ASINH(0.5)", 0.48121182505960347);
    assert_close("ASINH(-1E-8)", -1e-08);
    assert_close("ASINH(1E10)", 23.7189981105004);
    assert_close("ATANH(0.5)", 0.5493061443340549);
    assert_close("ATANH(-0.999)", -3.8002011672501994);
    assert_error("ACOSH(0.5)", ErrorValue::Num);
    assert_error("ATANH(1)", ErrorValue::Num);
    assert_error("ATANH(-1.5)", ErrorValue::Num);
    assert_error("ASINH(\"x\")", ErrorValue::Value);
}

#[test]
fn inverse_cotangents_match_documentation_and_oracles() {
    assert_published("ACOT(2)", 0.4636, 4);
    assert_published("ACOTH(6)", 0.168, 3);
    // mpmath 30 digits; ACOT of a negative number lies in (pi/2, pi)
    assert_close("ACOT(0.3)", 1.2793395323170296);
    assert_close("ACOT(-2)", 2.677945044588987);
    assert_close("ACOT(0)", std::f64::consts::FRAC_PI_2);
    assert_close("ACOTH(1.5)", 0.8047189562170501);
    assert_close("ACOTH(-1.0001)", -4.95176877564314);
    assert_close("ACOTH(1E8)", 1e-08);
    assert_error("ACOTH(0.5)", ErrorValue::Num);
    assert_error("ACOTH(1)", ErrorValue::Num);
    assert_error("ACOTH(-1)", ErrorValue::Num);
}

#[test]
fn reciprocal_trig_functions_match_documentation_and_oracles() {
    assert_published("COT(30)", -0.156, 3);
    assert_published("COT(45)", 0.617, 3);
    assert_published("COTH(2)", 1.037, 3);
    assert_published("CSC(15)", 1.538, 3);
    assert_published("CSCH(1.5)", 0.4696, 4);
    assert_published("SEC(45)", 1.90359, 5);
    assert_close_in(&[], "SECH(45)", 5.73e-20, 1e-3);
    // mpmath 30 digits
    assert_close("COT(0.7)", 1.1872418321266796);
    assert_close("COT(-3)", 7.015252551434534);
    assert_close("COTH(0.5)", 2.163953413738653);
    assert_close("COTH(-20)", -1.0);
    assert_close("CSC(1.2)", 1.0729163777098973);
    assert_close("CSC(-0.01)", -100.00166668611132);
    assert_close("CSCH(0.3)", 3.2838533966984236);
    assert_close("CSCH(-5)", -0.013476505830589087);
    assert_close("SEC(1)", 1.8508157176809257);
    assert_close("SEC(-2.5)", -1.2482156514688179);
    assert_close("SECH(0.25)", 0.9695436291402146);
    assert_close("SECH(30)", 1.871524593768035e-13);
    for name in ["COT", "COTH", "CSC", "CSCH"] {
        assert_error(&format!("{name}(0)"), ErrorValue::Div0);
    }
    assert_exact("SEC(0)", 1.0);
    assert_exact("SECH(0)", 1.0);
    for name in ["COT", "COTH", "CSC", "CSCH", "SEC", "SECH"] {
        assert_error(&format!("{name}(2^27)"), ErrorValue::Num);
        assert_error(&format!("{name}(-2^27)"), ErrorValue::Num);
        assert_error(&format!("{name}(\"x\")"), ErrorValue::Value);
        assert!(matches!(
            eval(&format!("{name}(2^27-1)")),
            CellValue::Number { .. }
        ));
    }
}

#[test]
fn roman_writes_every_form_excel_documents() {
    assert_text("ROMAN(499,0)", "CDXCIX");
    assert_text("ROMAN(499,1)", "LDVLIV");
    assert_text("ROMAN(499,2)", "XDIX");
    assert_text("ROMAN(499,3)", "VDIV");
    assert_text("ROMAN(499,4)", "ID");
    assert_text("ROMAN(2013,0)", "MMXIII");
    assert_text("ROMAN(499)", "CDXCIX");
    assert_text("ROMAN(499,)", "CDXCIX");
    assert_text("ROMAN(499,TRUE)", "CDXCIX");
    assert_text("ROMAN(499,FALSE)", "ID");
    assert_text("ROMAN(499,4.9)", "ID");
    assert_text("ROMAN(1999,0)", "MCMXCIX");
    assert_text("ROMAN(1999,1)", "MLMVLIV");
    assert_text("ROMAN(1999,2)", "MXMIX");
    assert_text("ROMAN(1999,3)", "MVMIV");
    assert_text("ROMAN(1999,4)", "MIM");
    assert_text("ROMAN(3999)", "MMMCMXCIX");
    assert_text("ROMAN(45,1)", "VL");
    assert_text("ROMAN(490,1)", "LDXL");
    assert_text("ROMAN(12.9)", "XII");
    assert_text("ROMAN(0)", "");
    assert_error("ROMAN(4000)", ErrorValue::Value);
    assert_error("ROMAN(-1)", ErrorValue::Value);
    assert_error("ROMAN(10,5)", ErrorValue::Value);
    assert_error("ROMAN(10,-1)", ErrorValue::Value);
    assert_error("ROMAN(\"x\")", ErrorValue::Value);
}

#[test]
fn arabic_reads_numerals_in_any_form() {
    assert_exact("ARABIC(\"LVII\")", 57.0);
    assert_exact("ARABIC(\"mcmxii\")", 1912.0);
    for form in 0..=4 {
        for value in [4, 9, 45, 49, 99, 490, 499, 999, 1999, 3999] {
            assert_exact(&format!("ARABIC(ROMAN({value},{form}))"), f64::from(value));
        }
    }
    assert_exact("ARABIC(\"  -xiv \")", -14.0);
    assert_exact("ARABIC(\"IIII\")", 4.0);
    assert_exact("ARABIC(\"\")", 0.0);
    assert_exact("ARABIC(REPT(\"M\",255))", 255_000.0);
    assert_eq!(
        eval_in(&[("A1", CellValue::Empty)], "ARABIC(A1)"),
        number(0.0)
    );
    assert_error("ARABIC(REPT(\"I\",256))", ErrorValue::Value);
    assert_error("ARABIC(12)", ErrorValue::Value);
    assert_error("ARABIC(TRUE)", ErrorValue::Value);
    assert_error("ARABIC(\"ABC\")", ErrorValue::Value);
    assert_error("ARABIC(1/0)", ErrorValue::Div0);
}

#[test]
fn base_and_decimal_convert_between_radixes() {
    assert_text("BASE(7,2)", "111");
    assert_text("BASE(100,16)", "64");
    assert_text("BASE(15,2,10)", "0000001111");
    assert_exact("DECIMAL(\"FF\",16)", 255.0);
    assert_exact("DECIMAL(111,2)", 7.0);
    assert_exact("DECIMAL(\"zap\",36)", 45_745.0);
    // exact python integer arithmetic
    assert_text("BASE(2^53-1,36)", "2GOSA7PA2GV");
    assert_text("BASE(2^53-1,2)", &"1".repeat(53));
    assert_text("BASE(123456789,7)", "3026236221");
    assert_text("BASE(0,2)", "0");
    assert_text("BASE(7.9,2.9)", "111");
    assert_exact("DECIMAL(\"zz9a\",36)", 1_678_654.0);
    assert_exact("DECIMAL(\"7fffffffff\",16)", 549_755_813_887.0);
    assert_exact("DECIMAL(\"\",2)", 0.0);
    assert_exact("DECIMAL(BASE(2^53-1,36),36)", 9007199254740991.0);
    assert_error("BASE(-1,2)", ErrorValue::Num);
    assert_error("BASE(2^53,2)", ErrorValue::Num);
    assert_error("BASE(7,1)", ErrorValue::Num);
    assert_error("BASE(7,37)", ErrorValue::Num);
    assert_error("BASE(7,2,256)", ErrorValue::Num);
    assert_error("BASE(7,2,-1)", ErrorValue::Num);
    assert_error("BASE(\"x\",2)", ErrorValue::Value);
    assert_error("DECIMAL(\"2\",2)", ErrorValue::Num);
    assert_error("DECIMAL(\"FF\",37)", ErrorValue::Num);
    assert_error("DECIMAL(\"FF\",1)", ErrorValue::Num);
    assert_error("DECIMAL(REPT(\"1\",256),2)", ErrorValue::Value);
    assert_error("DECIMAL(TRUE,36)", ErrorValue::Value);
}

#[test]
fn ceiling_and_floor_math_follow_their_modes() {
    assert_exact("CEILING.MATH(24.3,5)", 25.0);
    assert_exact("CEILING.MATH(6.7)", 7.0);
    assert_exact("CEILING.MATH(-8.1,2)", -8.0);
    assert_exact("CEILING.MATH(-5.5,2,-1)", -6.0);
    assert_exact("FLOOR.MATH(24.3,5)", 20.0);
    assert_exact("FLOOR.MATH(6.7)", 6.0);
    assert_exact("FLOOR.MATH(-8.1,2)", -10.0);
    assert_exact("FLOOR.MATH(-5.5,2,-1)", -4.0);
    // mpmath ceil and floor over the exact decimals
    assert_close("CEILING.MATH(-7.25,0.5,1)", -7.5);
    assert_close("CEILING.MATH(123.456,-0.05)", 123.5);
    assert_close("FLOOR.MATH(123.456,0.05)", 123.45);
    assert_close("FLOOR.MATH(-7.25,0.5,1)", -7.0);
    assert_exact("CEILING.MATH(5,0)", 0.0);
    assert_exact("FLOOR.MATH(-5,0)", 0.0);
    assert_exact("CEILING.MATH(-0.5)", 0.0);
    assert_error("CEILING.MATH(\"x\")", ErrorValue::Value);
    assert_error("FLOOR.MATH(1,\"x\")", ErrorValue::Value);
}

#[test]
fn precise_roundings_ignore_the_sign_of_significance() {
    for (number, significance, expected) in [
        ("4.3", "", 5.0),
        ("-4.3", "", -4.0),
        ("4.3", ",2", 6.0),
        ("4.3", ",-2", 6.0),
        ("-4.3", ",2", -4.0),
        ("-4.3", ",-2", -4.0),
    ] {
        assert_exact(
            &format!("CEILING.PRECISE({number}{significance})"),
            expected,
        );
        assert_exact(&format!("ISO.CEILING({number}{significance})"), expected);
    }
    assert_exact("FLOOR.PRECISE(-3.2,-1)", -4.0);
    assert_exact("FLOOR.PRECISE(3.2,1)", 3.0);
    assert_exact("FLOOR.PRECISE(-3.2,1)", -4.0);
    assert_exact("FLOOR.PRECISE(3.2,-1)", 3.0);
    assert_exact("FLOOR.PRECISE(3.2)", 3.0);
    // mpmath ceil and floor over the exact decimals
    assert_close("CEILING.PRECISE(-2.75,0.5)", -2.5);
    assert_close("FLOOR.PRECISE(17.3,4)", 16.0);
    assert_exact("CEILING.PRECISE(4.3,0)", 0.0);
    assert_error("FLOOR.PRECISE(\"x\")", ErrorValue::Value);
}

/// excel does not lose a multiple to binary noise: 0.3 / 0.1 is
/// 2.9999999999999996 in doubles, yet FLOOR.MATH(0.3, 0.1) is 0.3.
#[test]
fn significance_roundings_absorb_binary_noise_as_excel_does() {
    assert_exact("FLOOR.MATH(0.3,0.1)", 0.3);
    assert_exact("FLOOR.PRECISE(4.42,0.01)", 4.42);
    assert_exact("CEILING.MATH(0.3,0.1)", 0.3);
    assert_exact("CEILING.PRECISE(4.42,0.01)", 4.42);
    assert_exact("CEILING.MATH(1.000000000000001)", 2.0);
}

#[test]
fn combina_counts_combinations_with_repetition() {
    assert_exact("COMBINA(4,3)", 20.0);
    assert_exact("COMBINA(10,3)", 220.0);
    // mpmath 30 digits
    assert_exact("COMBINA(30,20)", 28_277_527_346_376.0);
    assert_exact("COMBINA(100,10)", 42_634_215_112_710.0);
    assert_exact("COMBINA(10.9,3.9)", 220.0);
    assert_exact("COMBINA(0,0)", 1.0);
    assert_exact("COMBINA(5,0)", 1.0);
    assert_error("COMBINA(-1,0)", ErrorValue::Num);
    assert_error("COMBINA(3,-1)", ErrorValue::Num);
    assert_error("COMBINA(3,5)", ErrorValue::Num);
    assert_error("COMBINA(1E10,1E10)", ErrorValue::Num);
    assert_error("COMBINA(\"x\",1)", ErrorValue::Value);
}

#[test]
fn factdouble_multiplies_every_second_integer() {
    assert_exact("FACTDOUBLE(6)", 48.0);
    assert_exact("FACTDOUBLE(7)", 105.0);
    // mpmath 30 digits
    assert_exact("FACTDOUBLE(25)", 7_905_853_580_625.0);
    assert_close("FACTDOUBLE(100)", 3.4243224702511973e+79);
    assert_exact("FACTDOUBLE(0)", 1.0);
    assert_exact("FACTDOUBLE(1)", 1.0);
    assert_exact("FACTDOUBLE(6.9)", 48.0);
    assert_error("FACTDOUBLE(-1)", ErrorValue::Num);
    assert_error("FACTDOUBLE(-0.5)", ErrorValue::Num);
    assert_error("FACTDOUBLE(400)", ErrorValue::Num);
}

#[test]
fn multinomial_divides_the_factorial_of_the_sum() {
    assert_exact("MULTINOMIAL(2,3,4)", 1260.0);
    // mpmath 30 digits
    assert_exact("MULTINOMIAL(5,7,9)", 232_792_560.0);
    assert_close("MULTINOMIAL(10,20,30)", 3.553261127084985e+24);
    let inputs = [
        ("A1", number(2.0)),
        ("A2", number(3.0)),
        ("A3", CellValue::Empty),
        ("A4", number(4.0)),
    ];
    assert_eq!(eval_in(&inputs, "MULTINOMIAL(A1:A4)"), number(1260.0));
    assert_exact("MULTINOMIAL(2.9,3.2,4.7)", 1260.0);
    assert_exact("MULTINOMIAL(7)", 1.0);
    assert_error("MULTINOMIAL(-1,2)", ErrorValue::Num);
    assert_error("MULTINOMIAL(1000,1000)", ErrorValue::Num);
    assert_error("MULTINOMIAL(\"x\",2)", ErrorValue::Value);
    assert_eq!(
        eval_in(
            &[("A1", number(2.0)), ("A2", text("x"))],
            "MULTINOMIAL(A1:A2)"
        ),
        error(ErrorValue::Value)
    );
}

#[test]
fn sqrtpi_takes_the_root_of_number_times_pi() {
    assert_published("SQRTPI(1)", 1.772453851, 9);
    assert_published("SQRTPI(2)", 2.506628275, 9);
    // mpmath 30 digits
    assert_close("SQRTPI(0.5)", 1.2533141373155003);
    assert_close("SQRTPI(1E6)", 1772.453850905516);
    assert_exact("SQRTPI(0)", 0.0);
    assert_error("SQRTPI(-1)", ErrorValue::Num);
}

#[test]
fn paired_square_sums_match_documentation_and_oracles() {
    let mut inputs = Vec::new();
    for (i, (x, y)) in [2.0, 3.0, 9.0, 1.0, 8.0, 7.0, 5.0]
        .iter()
        .zip([6.0, 5.0, 11.0, 7.0, 5.0, 4.0, 4.0])
        .enumerate()
    {
        inputs.push((format!("A{}", i + 2), number(*x)));
        inputs.push((format!("B{}", i + 2), number(y)));
    }
    let documented = |formula: &str| eval_with(&inputs, DateSystem::V1900, formula);
    assert_eq!(documented("SUMX2MY2(A2:A8,B2:B8)"), number(-55.0));
    assert_eq!(documented("SUMX2PY2(A2:A8,B2:B8)"), number(521.0));
    assert_eq!(documented("SUMXMY2(A2:A8,B2:B8)"), number(79.0));
    assert_exact("SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", -55.0);
    assert_exact("SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 521.0);
    assert_exact("SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 79.0);
    // mpmath 30 digits
    let xs = "{1.5,-2.25,3.1,0.4}";
    let ys = "{0.7,5.5,-1.2,2.05}";
    assert_close(&format!("SUMX2MY2({xs},{ys})"), -19.299999999999997);
    assert_close(&format!("SUMX2PY2({xs},{ys})"), 53.464999999999996);
    assert_close(&format!("SUMXMY2({xs},{ys})"), 81.915);
    let mixed = [
        ("A1", number(3.0)),
        ("A2", text("x")),
        ("A3", number(4.0)),
        ("B1", number(1.0)),
        ("B2", number(9.0)),
        ("B3", CellValue::Empty),
        ("A4", error(ErrorValue::Div0)),
    ];
    assert_eq!(eval_in(&mixed, "SUMX2PY2(A1:A3,B1:B3)"), number(10.0));
    assert_eq!(
        eval_in(&mixed, "SUMXMY2(A1:A4,B1:B4)"),
        error(ErrorValue::Div0)
    );
    assert_error("SUMX2MY2({1,2,3},{1,2})", ErrorValue::NA);
    assert_error("SUMXMY2({1,2},{1;2;3})", ErrorValue::NA);
}

#[test]
fn percentof_divides_the_subset_sum_by_the_total() {
    let inputs = [
        ("A1", number(10.0)),
        ("A2", number(20.0)),
        ("A3", number(30.0)),
        ("A4", number(40.0)),
        ("A5", text("n/a")),
        ("A6", CellValue::Bool { value: true }),
    ];
    // exact rational arithmetic: 30 / 100 and 70 / 100
    assert_close_in(&inputs, "PERCENTOF(A1:A2,A1:A6)", 0.3, 1e-12);
    assert_close_in(&inputs, "PERCENTOF(A3:A4,A1:A4)", 0.7, 1e-12);
    assert_close_in(&inputs, "_xlfn.PERCENTOF(A4,A1:A4)", 0.4, 1e-12);
    assert_eq!(
        eval_in(&inputs, "PERCENTOF(A1,A5:A6)"),
        error(ErrorValue::Div0)
    );
    assert_error("PERCENTOF(1)", ErrorValue::Value);
    assert_error("PERCENTOF(1/0,2)", ErrorValue::Div0);
}

#[test]
fn mdeterm_matches_documentation_and_oracles() {
    assert_close("MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})", 88.0);
    assert_close("MDETERM({3,6,1;1,1,0;3,10,2})", 1.0);
    assert_close("MDETERM({3,6;1,1})", -3.0);
    // mpmath 30 digits
    assert_close(
        "MDETERM({2.5,-1.25,0.75;1.5,3.125,-2;0.2,4.4,1.1})",
        37.6375,
    );
    assert_close(
        "MDETERM({0.3,7,-1.9,2;5.5,-0.6,1,3.3;-2.2,1.8,4,0.1;1,2,3,-4.75})",
        1236.1175,
    );
    assert_exact("MDETERM(4)", 4.0);
    assert_exact("MDETERM({1,2;2,4})", 0.0);
    let inputs = [
        ("A1", number(1.0)),
        ("B1", number(2.0)),
        ("A2", number(3.0)),
        ("B2", CellValue::Empty),
    ];
    assert_eq!(eval_in(&inputs, "MDETERM(A1:B2)"), error(ErrorValue::Value));
    assert_error("MDETERM({1,3,8,5;1,3,6,1})", ErrorValue::Value);
    assert_error("MDETERM({1,\"x\";3,4})", ErrorValue::Value);
    assert_error("MDETERM({1,2;3,4}/0)", ErrorValue::Div0);
}

/// excel eliminates with partial pivoting and returns the residue it leaves:
/// the singular 1..9 block comes back as 6.66134e-16, not 0.
#[test]
fn mdeterm_returns_excels_rounding_residue_for_a_singular_block() {
    assert_close("MDETERM({1,2,3;4,5,6;7,8,9})", 6.661338147750939e-16);
}

#[test]
fn minverse_matches_documentation_and_oracles() {
    assert_block("MINVERSE({4,-1;2,0})", 2, 2, &[0.0, 0.5, -1.0, 2.0], 0.0);
    assert_block(
        "MINVERSE({1,2,1;3,4,-1;0,2,0})",
        3,
        3,
        &[0.25, 0.25, -0.75, 0.0, 0.0, 0.5, 0.75, -0.25, -0.25],
        1e-12,
    );
    // mpmath 30 digits
    assert_block(
        "MINVERSE({2.5,-1.25,0.75;1.5,3.125,-2;0.2,4.4,1.1})",
        3,
        3,
        &[
            0.3251411491198937,
            0.12421122550647626,
            0.004151444702756559,
            -0.05446695450016606,
            0.06908003985386915,
            0.1627366323480571,
            0.15875124543341082,
            -0.29890401859847227,
            0.25738957157090664,
        ],
        1e-12,
    );
    assert_block("MINVERSE(4)", 1, 1, &[0.25], 0.0);
    assert_eq!(eval("INDEX(MINVERSE({4,-1;2,0}),2,2)"), number(2.0));
    assert_eq!(eval("SUM(MINVERSE({4,-1;2,0}))"), number(1.5));
    assert_eq!(block_error("MINVERSE({1,2;2,4})"), error(ErrorValue::Num));
    assert_eq!(block_error("MINVERSE({1,2,3})"), error(ErrorValue::Value));
    assert_eq!(
        block_error("MINVERSE({1,\"x\";3,4})"),
        error(ErrorValue::Value)
    );
}

#[test]
fn munit_builds_an_identity_block() {
    assert_block(
        "_xlfn.MUNIT(3)",
        3,
        3,
        &[1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0],
        0.0,
    );
    assert_block("MUNIT(2.9)", 2, 2, &[1.0, 0.0, 0.0, 1.0], 0.0);
    assert_block("MUNIT(1)", 1, 1, &[1.0], 0.0);
    assert_eq!(eval("SUM(MUNIT(4))"), number(4.0));
    assert_eq!(block_error("MUNIT(0)"), error(ErrorValue::Value));
    assert_eq!(block_error("MUNIT(-2)"), error(ErrorValue::Value));
    assert_eq!(block_error("MUNIT(\"x\")"), error(ErrorValue::Value));
}

#[test]
fn randarray_draws_from_the_seeded_stream() {
    let (rows, cols, draws) = block("_xlfn.RANDARRAY(3,2)", Some(42));
    assert_eq!((rows, cols), (3, 2));
    assert!(draws.iter().all(|x| (0.0..1.0).contains(x)));
    assert_eq!(block("RANDARRAY(3,2)", Some(42)).2, draws);
    let (_, _, first) = block("RAND()", Some(42));
    assert_eq!(first[0], draws[0]);
    let (_, _, bounded) = block("RANDARRAY(50,1,-3,2)", Some(7));
    assert!(bounded.iter().all(|x| (-3.0..2.0).contains(x)));
    let (_, _, dice) = block("RANDARRAY(4,5,1,6,TRUE)", Some(7));
    assert_eq!(dice.len(), 20);
    assert!(
        dice.iter()
            .all(|x| x.fract() == 0.0 && (1.0..=6.0).contains(x))
    );
    let (_, _, nudged) = block("RANDARRAY(30,1,0.5,2.5,TRUE)", Some(9));
    assert!(nudged.iter().all(|x| *x == 1.0 || *x == 2.0));
    assert_eq!(block("RANDARRAY()", Some(3)).0, 1);
    assert_eq!(block("RANDARRAY(,4)", Some(3)).1, 4);
    assert_eq!(block_error("RANDARRAY(0)"), error(ErrorValue::Calc));
    assert_eq!(block_error("RANDARRAY(2,0)"), error(ErrorValue::Calc));
    assert_eq!(block_error("RANDARRAY(-1)"), error(ErrorValue::Value));
    assert_eq!(block_error("RANDARRAY(2,2,5,1)"), error(ErrorValue::Value));
    assert_eq!(
        block_error("RANDARRAY(1,1,1.2,1.8,TRUE)"),
        error(ErrorValue::Value)
    );
}

#[test]
fn days360_matches_the_documented_examples() {
    assert_exact("DAYS360(DATE(2011,1,30),DATE(2011,2,1))", 1.0);
    assert_exact("DAYS360(DATE(2011,1,1),DATE(2011,12,31))", 360.0);
    assert_exact("DAYS360(DATE(2011,1,1),DATE(2011,2,1))", 30.0);
}

/// day counts worked by hand from the US and European rules microsoft
/// documents.
#[test]
fn days360_applies_the_us_and_european_rules() {
    let cases = [
        ((2011, 2, 28), (2011, 3, 31), 30.0, 32.0),
        ((2012, 2, 28), (2012, 3, 28), 30.0, 30.0),
        ((2012, 2, 29), (2012, 3, 28), 28.0, 29.0),
        ((2011, 1, 15), (2011, 2, 28), 43.0, 43.0),
        ((2011, 3, 15), (2011, 3, 31), 16.0, 15.0),
        ((2011, 3, 30), (2011, 3, 31), 0.0, 0.0),
        ((2011, 12, 31), (2011, 1, 1), -359.0, -359.0),
        ((2000, 1, 31), (2024, 2, 29), 8669.0, 8669.0),
    ];
    for ((y1, m1, d1), (y2, m2, d2), us, european) in cases {
        let dates = format!("DATE({y1},{m1},{d1}),DATE({y2},{m2},{d2})");
        assert_exact(&format!("DAYS360({dates})"), us);
        assert_exact(&format!("DAYS360({dates},FALSE)"), us);
        assert_exact(&format!("DAYS360({dates},TRUE)"), european);
        assert_exact(&format!("DAYS360({dates},1)"), european);
    }
    assert_exact("DAYS360(\"1/31/2011\",\"3/31/2011\",TRUE)", 60.0);
    assert_exact("DAYS360(DATE(2011,1,1)+0.9,DATE(2011,2,1))", 30.0);
    assert_error("DAYS360(-1,10)", ErrorValue::Num);
    assert_error("DAYS360(1,3000000)", ErrorValue::Num);
    assert_error("DAYS360(\"x\",10)", ErrorValue::Value);
    assert_error("DAYS360(1,10,\"x\")", ErrorValue::Value);
    assert_error("DAYS360(1)", ErrorValue::Value);
}

/// a 1904 workbook's serials are four years and a day behind the 1900 ones:
/// 39140 is 2011-02-28 there, a february month end, where on the 1900
/// calendar it would be 2007-02-27.
#[test]
fn days360_reads_1904_serials_on_their_own_calendar() {
    let inputs = cells(&[("A1", number(39_140.0)), ("B1", number(39_171.0))]);
    assert_eq!(
        eval_with(&inputs, DateSystem::V1904, "DAYS360(A1,B1)"),
        number(30.0)
    );
    assert_eq!(
        eval_with(&inputs, DateSystem::V1904, "DAYS360(A1,B1,TRUE)"),
        number(32.0)
    );
    assert_eq!(
        eval_with(&inputs, DateSystem::V1900, "DAYS360(A1,B1)"),
        number(33.0)
    );
    assert_eq!(
        eval_with(
            &[],
            DateSystem::V1904,
            "DAYS360(\"2/28/2011\",\"3/31/2011\")"
        ),
        number(30.0)
    );
    assert_eq!(
        eval_with(&inputs, DateSystem::V1904, "DAYS360(A1,\"3/31/2011\")"),
        number(30.0)
    );
    assert_eq!(
        eval_with(&inputs, DateSystem::V1904, "DAYS360(A1,2957003)"),
        eval_with(&[], DateSystem::V1900, "DAYS360(40602,2958465)")
    );
    assert_eq!(
        eval_with(&inputs, DateSystem::V1904, "DAYS360(A1,2957004)"),
        error(ErrorValue::Num)
    );
}

#[test]
fn stored_and_alias_names_share_one_implementation() {
    for (modern, other) in [
        ("CEILING.PRECISE(-4.3,2)", "ISO.CEILING(-4.3,2)"),
        ("CEILING.PRECISE(7.1,3)", "ISO.CEILING(7.1,3)"),
        ("COT(0.7)", "_xlfn.COT(0.7)"),
        ("ACOT(-2)", "_xlfn.ACOT(-2)"),
        ("CEILING.MATH(-5.5,2,-1)", "_xlfn.CEILING.MATH(-5.5,2,-1)"),
        ("FLOOR.MATH(-5.5,2,-1)", "_xlfn.FLOOR.MATH(-5.5,2,-1)"),
        ("COMBINA(4,3)", "_xlfn.COMBINA(4,3)"),
        ("ARABIC(\"LVII\")", "_xlfn.ARABIC(\"LVII\")"),
        ("BASE(100,16)", "_xlfn.BASE(100,16)"),
        ("DECIMAL(\"FF\",16)", "_xlfn.DECIMAL(\"FF\",16)"),
        ("SEC(1)", "_xlfn.SEC(1)"),
        ("SECH(1)", "_xlfn.SECH(1)"),
        ("CSC(1)", "_xlfn.CSC(1)"),
        ("CSCH(1)", "_xlfn.CSCH(1)"),
        ("COTH(1)", "_xlfn.COTH(1)"),
        ("ACOTH(2)", "_xlfn.ACOTH(2)"),
    ] {
        assert_eq!(eval(modern), eval(other), "{modern} vs {other}");
        assert!(matches!(
            eval(modern),
            CellValue::Number { .. } | CellValue::Text { .. }
        ));
    }
}

#[test]
fn scalar_functions_lift_over_array_arguments() {
    assert_block(
        "COT({0.7,-3})",
        1,
        2,
        &[1.1872418321266793, 7.015252551434533],
        1e-12,
    );
    assert_block("FACTDOUBLE({6;7})", 2, 1, &[48.0, 105.0], 0.0);
    assert_block(
        "DAYS360(DATE(2011,1,1),DATE(2011,{2,3},1))",
        1,
        2,
        &[30.0, 60.0],
        0.0,
    );
}

/// (1 + 10^16)! / (1! 10^16!) is 10^16 + 1 whichever order the counts come
/// in; the 1 must not vanish into a rounded running total.
#[test]
fn multinomial_keeps_a_small_count_beside_a_huge_one() {
    assert_close("MULTINOMIAL(1,1E16)", 1e16);
    assert_close("MULTINOMIAL(1E16,1)", 1e16);
    // exact integer arithmetic: (10^16 + 2)(10^16 + 1) / 2
    assert_close("MULTINOMIAL(2,1E16)", 5.000000000000001e31);
    assert_close("MULTINOMIAL(1E16,2)", 5.000000000000001e31);
    assert_close("MULTINOMIAL(2^53,1)", 9_007_199_254_740_992.0);
    assert_close("MULTINOMIAL(1,2^53)", 9_007_199_254_740_992.0);
    assert_exact("MULTINOMIAL(0,1E16)", 1.0);
    assert_error("MULTINOMIAL(-1E-200,2)", ErrorValue::Num);
}

/// 10^15 + 0.75 has a real fraction: FLOOR.MATH must round it down to
/// 10^15, as microsoft documents, and the ceilings up to 10^15 + 1.
#[test]
fn significance_roundings_keep_real_fractions_of_large_numbers() {
    assert_exact("FLOOR.MATH(1000000000000000+0.75)", 1e15);
    assert_exact("FLOOR.PRECISE(1000000000000000+0.75)", 1e15);
    assert_exact("FLOOR.MATH(1000000000000000+0.875)", 1e15);
    assert_exact("CEILING.MATH(1000000000000000+0.75)", 1e15 + 1.0);
    assert_exact("CEILING.PRECISE(1000000000000000+0.125)", 1e15 + 1.0);
    assert_exact("ISO.CEILING(1000000000000000+0.25)", 1e15 + 1.0);
    assert_exact("FLOOR.MATH(-1000000000000000-0.75)", -1e15 - 1.0);
    assert_exact("FLOOR.MATH(-1000000000000000-0.75,1,1)", -1e15);
    assert_exact("CEILING.MATH(-1000000000000000-0.75)", -1e15);
    assert_exact("CEILING.MATH(-1000000000000000-0.75,1,1)", -1e15 - 1.0);
    assert_exact("FLOOR.MATH(1000000000000000+0.75,0.5)", 1e15 + 0.5);
    assert_exact("CEILING.MATH(1000000000000000+0.75,0.5)", 1e15 + 1.0);
    assert_exact("FLOOR.MATH(2^52-0.5)", 4_503_599_627_370_495.0);
    assert_exact("CEILING.MATH(2^51+0.5)", 2_251_799_813_685_249.0);
    assert_exact("FLOOR.MATH(2^53+2,0.5)", 9_007_199_254_740_994.0);
    assert_exact("CEILING.MATH(1E19)", 1e19);
}

/// a multiple that exists must come back even when the quotient overflows or
/// underflows on the way.
#[test]
fn significance_roundings_survive_extreme_magnitudes() {
    assert_exact("CEILING.MATH(1E308,0.1)", 1e308);
    assert_exact("FLOOR.PRECISE(1E300,1E-300)", 1e300);
    assert_exact("CEILING.MATH(1E-200,1E200)", 1e200);
    assert_exact("FLOOR.MATH(1E-200,1E200)", 0.0);
    assert_exact("FLOOR.MATH(-1E-200,1E200)", -1e200);
    assert_exact("CEILING.PRECISE(-1E-200,1E200)", 0.0);
    assert_exact("CEILING.MATH(0.5,1E-308)", 0.5);
    assert_error("CEILING.MATH(1.7E308,1E308)", ErrorValue::Num);
}

#[test]
fn reciprocal_hyperbolics_do_not_underflow_to_zero_early() {
    // mpmath 40 digits; the results are subnormal, where a double keeps
    // about 11 significant digits
    assert_close_in(&[], "SECH(720)", 4.06446160484e-313, 1e-10);
    assert_close_in(&[], "SECH(-720)", 4.06446160484e-313, 1e-10);
    assert_close_in(&[], "CSCH(-730)", -1.8452626e-317, 1e-6);
    assert_exact("SECH(800)", 0.0);
    assert_close("CSCH(1E-200)", 1e200);
    assert_error("CSCH(1E-309)", ErrorValue::Num);
    assert_error("COT(1E-310)", ErrorValue::Num);
}

#[test]
fn sqrtpi_and_paired_sums_survive_extreme_magnitudes() {
    // mpmath 40 digits
    assert_close("SQRTPI(1E308)", 1.772_453_850_905_516_2e154);
    assert_close("SQRTPI(5E-324)", 3.939_737_305_158_755e-162);
    assert_exact("SUMX2MY2(1E200,1E200)", 0.0);
    assert_exact("SUMX2MY2(1E308,-1E308)", 0.0);
    assert_error("SUMX2MY2(1.7E308,1E308)", ErrorValue::Num);
    assert_error("SUMXMY2(1E308,-1E308)", ErrorValue::Num);
}

#[test]
fn mdeterm_survives_an_intermediate_overflow() {
    assert_close("MDETERM({1E200,0,0;0,1E200,0;0,0,1E-200})", 1e200);
    assert_close("MDETERM({1E-200,0,0;0,1E-200,0;0,0,1E200})", 1e-200);
    assert_error("MDETERM({1E200,0;0,1E200})", ErrorValue::Num);
    assert_exact("MDETERM({1E-200,0;0,1E-200})", 0.0);
}

#[test]
fn negative_fractions_are_refused_before_truncation() {
    assert_error("BASE(-0.5,2)", ErrorValue::Num);
    assert_error("BASE(5,2,-0.5)", ErrorValue::Num);
    assert_error("ROMAN(5,-0.5)", ErrorValue::Value);
    assert_error("ROMAN(-0.5)", ErrorValue::Value);
    assert_error("DAYS360(-0.5,10)", ErrorValue::Num);
    assert_error("COMBINA(-1E-200,0)", ErrorValue::Num);
    assert_error("FACTDOUBLE(-1E-200)", ErrorValue::Num);
    assert_eq!(block_error("RANDARRAY(-0.5)"), error(ErrorValue::Value));
    assert_eq!(block_error("MUNIT(-0.5)"), error(ErrorValue::Value));
}

#[test]
fn whole_number_limits_hold_past_2_to_the_53() {
    assert!(matches!(eval("FACTDOUBLE(300)"), CellValue::Number { .. }));
    assert_error("FACTDOUBLE(301)", ErrorValue::Num);
    assert_error("FACTDOUBLE(2^53+1)", ErrorValue::Num);
    assert_error("FACTDOUBLE(1E300)", ErrorValue::Num);
    assert_exact("COMBINA(1E16,1)", 1e16);
    assert_exact("COMBINA(2^53+2,1)", 9_007_199_254_740_994.0);
    assert_error("COMBINA(1E300,2)", ErrorValue::Num);
    assert_error("MULTINOMIAL(1E300,1E300)", ErrorValue::Num);
    assert_text("BASE(2^52-0.5,2)", &"1".repeat(52));
    assert_error("ROMAN(1E200)", ErrorValue::Value);
    assert_error("DAYS360(1,2^53)", ErrorValue::Num);
    assert_eq!(block_error("MUNIT(1E200)"), error(ErrorValue::Value));
    assert_eq!(block_error("MUNIT(65536)"), error(ErrorValue::Num));
    assert_eq!(
        block_error("RANDARRAY(1048576,1048576)"),
        error(ErrorValue::Num)
    );
    assert_eq!(block_error("RANDARRAY(1E300)"), error(ErrorValue::Value));
}

/// bounds that span the whole double range still draw finite numbers.
#[test]
fn randarray_draws_across_the_whole_double_range() {
    let (_, _, wide) = block("RANDARRAY(20,1,-1E308,1E308)", Some(5));
    assert!(
        wide.iter()
            .all(|x| x.is_finite() && (-1e308..=1e308).contains(x))
    );
    let (_, _, whole) = block("RANDARRAY(20,1,-1E308,1E308,TRUE)", Some(5));
    assert!(whole.iter().all(|x| x.is_finite() && x.fract() == 0.0));
    let (_, _, big) = block("RANDARRAY(20,1,1E16,1E16+4,TRUE)", Some(5));
    assert!(big.iter().all(|x| (1e16..=1e16 + 4.0).contains(x)));
    assert_eq!(
        block_error("RANDARRAY(1,1,1000000000000000+0.25,1000000000000000+0.75,TRUE)"),
        error(ErrorValue::Value)
    );
}
