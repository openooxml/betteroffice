//! cash-flow financial functions, checked against microsoft's documented
//! examples, 30-digit mpmath oracles and values excel cached in real
//! workbooks.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, Workbook};

const RESULT: &str = "ZZ1";

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn column_name(mut index: usize) -> String {
    let mut name = Vec::new();
    loop {
        name.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    name.reverse();
    String::from_utf8(name).unwrap()
}

/// numbers down column `col` from `row`.
fn down(col: &str, row: u32, values: &[f64]) -> Vec<(String, CellValue)> {
    values
        .iter()
        .enumerate()
        .map(|(i, v)| (format!("{col}{}", row + i as u32), number(*v)))
        .collect()
}

/// numbers across row `row` from column A.
fn across(row: u32, values: &[f64]) -> Vec<(String, CellValue)> {
    values
        .iter()
        .enumerate()
        .map(|(i, v)| (format!("{}{row}", column_name(i)), number(*v)))
        .collect()
}

fn eval_in(inputs: &[(String, CellValue)], formula: &str) -> CellValue {
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

fn value_in(inputs: &[(String, CellValue)], formula: &str) -> f64 {
    match eval_in(inputs, formula) {
        CellValue::Number { value } => value,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

/// within `rel` of `want`, relative to its magnitude.
fn close_in(inputs: &[(String, CellValue)], formula: &str, want: f64, rel: f64) {
    let got = value_in(inputs, formula);
    let scale = want.abs().max(f64::MIN_POSITIVE);
    assert!(
        (got - want).abs() <= rel * scale,
        "{formula}: {got:e} vs {want:e} (rel {:e})",
        (got - want).abs() / scale
    );
}

fn close(formula: &str, want: f64, rel: f64) {
    close_in(&[], formula, want, rel);
}

/// within `abs` of `want`: microsoft's examples print rounded results.
fn near_in(inputs: &[(String, CellValue)], formula: &str, want: f64, abs: f64) {
    let got = value_in(inputs, formula);
    assert!((got - want).abs() <= abs, "{formula}: {got:e} vs {want:e}");
}

fn near(formula: &str, want: f64, abs: f64) {
    near_in(&[], formula, want, abs);
}

fn fails_in(inputs: &[(String, CellValue)], formula: &str, want: ErrorValue) {
    assert_eq!(
        eval_in(inputs, formula),
        CellValue::Error { value: want },
        "{formula}"
    );
}

fn fails(formula: &str, want: ErrorValue) {
    fails_in(&[], formula, want);
}

#[test]
fn rate_microsoft_example() {
    let inputs = down("A", 2, &[4.0, -200.0, 8000.0]);
    near_in(&inputs, "RATE(A2*12,A3,A4)", 0.0077, 5e-5);
    near_in(&inputs, "RATE(A2*12,A3,A4)*12", 0.0924, 5e-5);
}

/// excel accepts a rate once the step and the balance left are both under
/// 1e-7, so its answer is an iterative inverse good to about 1e-10.
#[test]
fn rate_matches_mpmath_roots() {
    // mpmath 30 digits
    close("RATE(48,-200,8000)", 0.007_701_472_488_202_044, 1e-9);
    close("RATE(360,-1000,150000)", 0.005_850_253_376_759_662, 1e-9);
    close("RATE(10,-100,800,0,1)", 0.053_446_167_393_037_78, 1e-9);
    close("RATE(5,0,-1000,1500)", 0.084_471_771_197_698_61, 1e-9);
    close("RATE(5,,-1000,1500,,0.5)", 0.084_471_771_197_698_61, 1e-9);
}

/// growth rates excel cached for `RATE(2025-2010,,-start,end)`, the idiom of
/// real forecasting workbooks, where the payment is left out.
#[test]
fn rate_matches_excel_cached_growth_rates() {
    for (years, start, end, excel) in [
        ("2025-2010", 144.0, 3.0, -0.227_466_625_103_029_95),
        ("2025-2010", 99.0, 19.0, -0.104_206_526_944_193_05),
        ("2026-2025", 4.0, 5.0, 0.249_999_999_999_999_92),
        ("2026-2025", 19.0, 7.0, -0.631_578_947_368_421),
        ("2025-2010", 16_675.223, 23_712.237, 0.023_748_767_146_439_4),
        (
            "2036-2026",
            24_215.977_5,
            28_947.711_229_424_3,
            0.018_008_074_563_340_35,
        ),
    ] {
        let inputs = [
            ("C10".to_string(), number(start)),
            ("C25".to_string(), number(end)),
        ];
        close_in(&inputs, &format!("RATE({years},,-C10,C25)"), excel, 1e-12);
    }
}

/// excel's secant search stops on a step under 1e-7, so a flat series gets
/// the residue of its last step rather than zero; excel cached these.
#[test]
fn rate_of_a_flat_series_is_excels_secant_residue() {
    let inputs = [
        ("C10".to_string(), number(7.0)),
        ("C25".to_string(), number(7.0)),
    ];
    near_in(
        &inputs,
        "RATE(2046-2026,,-C10,C25)",
        3.524_632_992_385_166e-12,
        1e-15,
    );
    near_in(
        &inputs,
        "RATE(2036-2026,,-C10,C25)",
        6.730_659_157_425_639e-16,
        1e-15,
    );
}

/// with nothing left at the end the balance only reaches zero at -100%;
/// excel answers the 200th secant step, whose balance is already negligible.
#[test]
fn rate_answers_its_last_step_once_the_balance_is_negligible() {
    let inputs = [
        ("C10".to_string(), number(49.0)),
        ("C25".to_string(), number(0.0)),
    ];
    close_in(
        &inputs,
        "RATE(2046-2026,,-C10,C25)",
        -0.999_112_071_912_539_3,
        1e-12,
    );
}

#[test]
fn rate_errors() {
    fails("RATE(10,100,1000)", ErrorValue::Num);
    fails("RATE(0,0,-100,100)", ErrorValue::Num);
    fails("RATE(10,-100)", ErrorValue::Value);
    fails("RATE(10,-100,800,0,0,0.1,1)", ErrorValue::Value);
    fails("RATE(\"ten\",-100,800)", ErrorValue::Value);
    fails("RATE(10,-100,1/0)", ErrorValue::Div0);
}

#[test]
fn irr_microsoft_examples_and_oracles() {
    let inputs = down(
        "A",
        2,
        &[-70_000.0, 12_000.0, 15_000.0, 18_000.0, 21_000.0, 26_000.0],
    );
    near_in(&inputs, "IRR(A2:A6)", -0.021, 5e-4);
    near_in(&inputs, "IRR(A2:A7)", 0.087, 5e-4);
    near_in(&inputs, "IRR(A2:A4,-10%)", -0.444, 5e-4);
    // mpmath 30 digits
    close_in(&inputs, "IRR(A2:A6)", -0.021_244_848_273_410_992, 1e-12);
    close_in(&inputs, "IRR(A2:A7)", 0.086_630_948_036_531_61, 1e-12);
    close_in(&inputs, "IRR(A2:A4,-10%)", -0.443_506_941_334_740_56, 1e-12);
    close("IRR({-100,39,59,55,20})", 0.280_948_421_159_961_1, 1e-12);
}

#[test]
fn irr_skips_text_and_blanks_in_ranges() {
    let mut inputs = down("A", 1, &[-100.0, 39.0, 59.0, 55.0, 20.0]);
    inputs.push(("A6".to_string(), text("total")));
    close_in(&inputs, "IRR(A1:A8)", 0.280_948_421_159_961_1, 1e-12);
}

#[test]
fn irr_errors() {
    let inputs = down("A", 1, &[100.0, 200.0, 300.0]);
    fails_in(&inputs, "IRR(A1:A3)", ErrorValue::Num);
    fails("IRR({-1,-2,-3})", ErrorValue::Num);
    let with_error = [
        ("A1".to_string(), number(-100.0)),
        (
            "A2".to_string(),
            CellValue::Error {
                value: ErrorValue::Div0,
            },
        ),
        ("A3".to_string(), number(50.0)),
    ];
    fails_in(&with_error, "IRR(A1:A3)", ErrorValue::Div0);
    fails("IRR()", ErrorValue::Value);
}

#[test]
fn mirr_microsoft_examples_and_oracles() {
    let inputs = down(
        "A",
        2,
        &[-120_000.0, 39_000.0, 30_000.0, 21_000.0, 37_000.0, 46_000.0],
    );
    near_in(&inputs, "MIRR(A2:A7,10%,12%)", 0.126, 5e-4);
    near_in(&inputs, "MIRR(A2:A5,10%,12%)", -0.048, 5e-4);
    near_in(&inputs, "MIRR(A2:A7,10%,14%)", 0.135, 5e-4);
    // mpmath 30 digits
    close_in(
        &inputs,
        "MIRR(A2:A7,10%,12%)",
        0.126_094_130_365_905_14,
        1e-12,
    );
    close_in(
        &inputs,
        "MIRR(A2:A5,10%,12%)",
        -0.048_044_655_249_980_82,
        1e-12,
    );
    close_in(
        &inputs,
        "MIRR(A2:A7,10%,14%)",
        0.134_759_110_828_314_85,
        1e-12,
    );
}

#[test]
fn mirr_needs_an_inflow_and_an_outflow() {
    fails("MIRR({100,200,300},0.1,0.1)", ErrorValue::Div0);
    fails("MIRR({-100,-200},0.1,0.1)", ErrorValue::Div0);
    fails("MIRR({-100,200},0.1)", ErrorValue::Value);
}

fn xirr_example() -> Vec<(String, CellValue)> {
    let mut inputs = down("A", 1, &[-10_000.0, 2_750.0, 4_250.0, 3_250.0, 2_750.0]);
    inputs.extend(down(
        "B",
        1,
        &[39_448.0, 39_508.0, 39_751.0, 39_859.0, 39_904.0],
    ));
    inputs
}

#[test]
fn xnpv_microsoft_example_and_oracles() {
    let inputs = xirr_example();
    near_in(&inputs, "XNPV(0.09,A1:A5,B1:B5)", 2_086.65, 5e-3);
    // mpmath 30 digits
    close_in(
        &inputs,
        "XNPV(0.09,A1:A5,B1:B5)",
        2_086.647_602_031_536_7,
        1e-12,
    );
    close_in(
        &inputs,
        "XNPV(-0.5,A1:A5,B1:B5)",
        14_268.654_346_155_818,
        1e-12,
    );
}

#[test]
fn xnpv_truncates_dates_and_takes_them_in_any_order() {
    let mut inputs = xirr_example();
    inputs.push(("B3".to_string(), number(39_751.75)));
    close_in(
        &inputs,
        "XNPV(0.09,A1:A5,B1:B5)",
        2_086.647_602_031_536_7,
        1e-12,
    );
    let shuffled = [
        down("A", 1, &[-10_000.0, 2_750.0, 3_250.0, 4_250.0, 2_750.0]),
        down("B", 1, &[39_448.0, 39_904.0, 39_859.0, 39_751.0, 39_508.0]),
    ]
    .concat();
    close_in(
        &shuffled,
        "XNPV(0.09,A1:A5,B1:B5)",
        2_086.647_602_031_536_7,
        1e-12,
    );
}

#[test]
fn xnpv_errors() {
    let mut inputs = xirr_example();
    fails_in(&inputs, "XNPV(0.09,A1:A5,B1:B4)", ErrorValue::Num);
    fails_in(&inputs, "XNPV(0.09,A1:A5)", ErrorValue::Value);
    inputs.push(("B4".to_string(), number(39_000.0)));
    fails_in(&inputs, "XNPV(0.09,A1:A5,B1:B5)", ErrorValue::Num);
    inputs.push(("B4".to_string(), text("soon")));
    fails_in(&inputs, "XNPV(0.09,A1:A5,B1:B5)", ErrorValue::Value);
    inputs.push(("B4".to_string(), number(-5.0)));
    fails_in(&inputs, "XNPV(0.09,A1:A5,B1:B5)", ErrorValue::Value);
}

/// microsoft prints 0.373362535; excel's answer is only good to the 1e-8 it
/// iterates to, and the exact root is 0.3733625335.
#[test]
fn xirr_microsoft_example_and_oracle() {
    let inputs = xirr_example();
    near_in(&inputs, "XIRR(A1:A5,B1:B5,0.1)", 0.373_362_535, 1e-8);
    // mpmath 30 digits
    close_in(
        &inputs,
        "XIRR(A1:A5,B1:B5)",
        0.373_362_533_518_831_53,
        1e-12,
    );
}

/// half-yearly dates and four cash-flow rows from a real appraisal workbook,
/// with the results excel cached. excel stops iterating once its estimate
/// moves by less than 1e-8, so its values sit up to about 4e-8 (relative)
/// from the exact root; the mpmath roots hold to 1e-12.
#[test]
fn xirr_matches_excel_cached_appraisals() {
    let dates = [
        41_547.0, 41_729.0, 41_912.0, 42_094.0, 42_277.0, 42_460.0, 42_643.0, 42_825.0, 43_008.0,
        43_190.0, 43_373.0, 43_555.0, 43_738.0, 43_921.0, 44_104.0, 44_286.0, 44_469.0, 44_651.0,
        44_834.0, 45_016.0, 45_199.0, 45_382.0, 45_565.0, 45_747.0, 45_930.0, 46_112.0, 46_295.0,
        46_477.0, 46_660.0, 46_843.0, 47_026.0, 47_208.0, 47_391.0, 47_573.0, 47_756.0, 47_938.0,
        48_121.0, 48_304.0, 48_487.0, 48_669.0, 48_852.0, 49_034.0, 49_217.0, 49_399.0, 49_582.0,
        49_765.0,
    ];
    let tail = [
        585.0, 580.0, 575.0, 570.0, 707.0, 694.0, 681.0, 668.0, 655.0, 642.0, 629.0, 616.0, 603.0,
        590.0, 577.0, 564.0, 551.0, 586.0, 571.0, 606.0, 591.0, 576.0, 701.0, 726.0, 701.0, 701.0,
        676.0, 676.0, 651.0, 1_026.0, 1_001.0, 1_000.0, 1_100.0,
    ];
    let full = |head: [f64; 6], tail: &[f64]| {
        let mut row = head.to_vec();
        row.extend_from_slice(tail);
        row.resize(dates.len(), 0.0);
        row
    };
    let first = [-10_100.0, 500.0, 500.0, 500.0, 595.0, 590.0];
    let mut second_tail = tail;
    second_tail[..4].copy_from_slice(&[645.0, 580.0, 635.0, 570.0]);
    let second = [-10_040.0, 500.0, 560.0, 500.0, 655.0, 590.0];
    for (flows, excel, exact) in [
        (
            full(first, &tail),
            0.106_996_637_582_778_94,
            0.106_996_638_671_615_2,
        ),
        (
            full(first, &[]),
            -0.534_025_494_009_256_4,
            -0.534_025_492_126_929_7,
        ),
        (
            full(second, &second_tail),
            0.110_701_328_516_006_45,
            0.110_701_330_640_124_98,
        ),
        (
            full(second, &[]),
            -0.522_735_539_823_770_7,
            -0.522_735_537_029_078_5,
        ),
    ] {
        let inputs = [across(4, &dates), across(87, &flows)].concat();
        close_in(&inputs, "XIRR(A87:AT87,A$4:AT$4)", excel, 5e-8);
        // mpmath 30 digits
        close_in(&inputs, "XIRR(A87:AT87,A$4:AT$4)", exact, 1e-12);
    }
}

/// all-zero rows in the same workbook: excel's XIRR is `#NUM!`, which the
/// workbook's `IF(ISERROR(..),0,..)` guard turns into the cached zero.
#[test]
fn xirr_without_both_signs_is_num() {
    let inputs = [
        across(4, &[694_052.0, 694_235.0, 694_417.0]),
        across(87, &[0.0, 0.0, 0.0]),
    ]
    .concat();
    fails_in(&inputs, "XIRR(A87:C87,A4:C4)", ErrorValue::Num);
    assert_eq!(
        eval_in(
            &inputs,
            "IF(ISERROR(XIRR(A87:C87,A$4:C$4)),0,XIRR(A87:C87,A$4:C$4))"
        ),
        number(0.0)
    );
}

#[test]
fn xirr_errors() {
    let mut inputs = xirr_example();
    fails_in(&inputs, "XIRR(A1:A5,B1:B4)", ErrorValue::Num);
    fails_in(&inputs, "XIRR(A1:A5)", ErrorValue::Value);
    inputs.push(("B5".to_string(), number(39_000.0)));
    fails_in(&inputs, "XIRR(A1:A5,B1:B5)", ErrorValue::Num);
    inputs.push(("B5".to_string(), text("later")));
    fails_in(&inputs, "XIRR(A1:A5,B1:B5)", ErrorValue::Value);
    inputs.push(("B5".to_string(), number(39_904.0)));
    inputs.push(("A2".to_string(), text("n/a")));
    fails_in(&inputs, "XIRR(A1:A5,B1:B5)", ErrorValue::Value);
}

#[test]
fn ipmt_microsoft_examples_and_oracles() {
    near("IPMT(0.1/12,1,3*12,8000)", -66.67, 5e-3);
    near("IPMT(0.1,3,3,8000)", -292.45, 5e-3);
    // mpmath 30 digits
    close("IPMT(0.1/12,1,3*12,8000)", -66.666_666_666_666_67, 1e-12);
    close("IPMT(0.1,3,3,8000)", -292.447_129_909_365_6, 1e-12);
    close(
        "IPMT(0.05/12,7,60,20000,1000,1)",
        -75.219_158_756_315_98,
        1e-12,
    );
}

#[test]
fn ipmt_of_the_first_payment_due_at_the_start_is_zero() {
    assert_eq!(eval("IPMT(0.05/12,1,60,20000,0,1)"), number(0.0));
}

#[test]
fn ppmt_microsoft_examples_and_oracles() {
    near("PPMT(0.1/12,1,2*12,2000)", -75.62, 5e-3);
    near("PPMT(0.08,10,10,200000)", -27_598.05, 5e-3);
    // mpmath 30 digits
    close("PPMT(0.1/12,1,2*12,2000)", -75.623_186_008_366_35, 1e-12);
    close("PPMT(0.08,10,10,200000)", -27_598.053_462_421_376, 1e-12);
    close(
        "PPMT(0.05/12,7,60,20000,1000,1)",
        -315.282_988_819_680_3,
        1e-12,
    );
}

#[test]
fn payment_split_errors() {
    for name in ["IPMT", "PPMT"] {
        fails(&format!("{name}(0.1,0,3,8000)"), ErrorValue::Num);
        fails(&format!("{name}(0.1,4,3,8000)"), ErrorValue::Num);
        fails(&format!("{name}(0.1,1,3)"), ErrorValue::Value);
        fails(&format!("{name}(\"x\",1,3,8000)"), ErrorValue::Value);
    }
}

#[test]
fn cumipmt_and_cumprinc_microsoft_examples_and_oracles() {
    near(
        "CUMIPMT(0.09/12,30*12,125000,13,24,0)",
        -11_135.232_13,
        5e-6,
    );
    near("CUMIPMT(0.09/12,30*12,125000,1,1,0)", -937.5, 5e-3);
    near(
        "CUMPRINC(0.09/12,30*12,125000,13,24,0)",
        -934.107_123_4,
        5e-8,
    );
    near("CUMPRINC(0.09/12,30*12,125000,1,1,0)", -68.278_271_18, 5e-9);
    // mpmath 30 digits
    close(
        "CUMIPMT(0.09/12,360,125000,13,24,0)",
        -11_135.232_130_750_843,
        1e-12,
    );
    close(
        "CUMPRINC(0.09/12,360,125000,13,24,0)",
        -934.107_123_420_898_3,
        1e-12,
    );
    close(
        "CUMIPMT(0.06/12,120,50000,1,12,1)",
        -2_634.446_843_646_35,
        1e-12,
    );
    close(
        "CUMPRINC(0.06/12,120,50000,1,12,1)",
        -3_993.642_824_511_827_3,
        1e-12,
    );
    close(
        "CUMIPMT(0.06/12,120,50000,1.9,12.7,1)",
        -2_634.446_843_646_35,
        1e-12,
    );
}

#[test]
fn cumulative_errors() {
    for name in ["CUMIPMT", "CUMPRINC"] {
        for args in [
            "0,360,125000,1,1,0",
            "0.01,0,125000,1,1,0",
            "0.01,360,0,1,1,0",
            "0.01,360,125000,0,1,0",
            "0.01,360,125000,5,4,0",
            "0.01,360,125000,1,361,0",
            "0.01,360,125000,1,1,2",
            "0.01,360,125000,1,1,0.5",
        ] {
            fails(&format!("{name}({args})"), ErrorValue::Num);
        }
        fails(&format!("{name}(0.01,360,125000,1,1)"), ErrorValue::Value);
    }
}

#[test]
fn ispmt_microsoft_examples() {
    near("ISPMT(0.1/12,1,3*12,8000000)", -64_814.8, 5e-2);
    near("ISPMT(0.1,1,3,8000000)", -533_333.0, 0.5);
    // mpmath 30 digits
    close(
        "ISPMT(0.1/12,1,3*12,8000000)",
        -64_814.814_814_814_82,
        1e-12,
    );
    fails("ISPMT(0.1,1,0,8000000)", ErrorValue::Div0);
    fails("ISPMT(0.1,1,3)", ErrorValue::Value);
}

#[test]
fn sln_and_syd_microsoft_examples() {
    assert_eq!(eval("SLN(30000,7500,10)"), number(2_250.0));
    near("SYD(30000,7500,10,1)", 4_090.91, 5e-3);
    near("SYD(30000,7500,10,10)", 409.09, 5e-3);
    // mpmath 30 digits
    close("SYD(30000,7500,10,1)", 4_090.909_090_909_091, 1e-12);
    close("SYD(1000,100,4.5,2)", 254.545_454_545_454_53, 1e-12);
    fails("SLN(30000,7500,0)", ErrorValue::Div0);
    fails("SYD(30000,7500,10,11)", ErrorValue::Num);
    fails("SYD(30000,7500,0,1)", ErrorValue::Num);
    fails("SYD(30000,7500,10,0)", ErrorValue::Num);
    fails("SLN(30000,7500)", ErrorValue::Value);
}

#[test]
fn db_microsoft_example() {
    // mpmath 30 digits, matching the documented 186,083.33 .. 15,845.10
    for (period, want) in [
        (1, 186_083.333_333_333_34),
        (2, 259_639.416_666_666_66),
        (3, 176_814.442_75),
        (4, 120_410.635_512_75),
        (5, 81_999.642_784_182_75),
        (6, 55_841.756_736_028_45),
        (7, 15_845.098_473_848_073),
    ] {
        close(&format!("DB(1000000,100000,6,{period},7)"), want, 1e-12);
    }
    close("DB(5000,500,4,3)", 691.698_36, 1e-12);
    close("DB(5000,500,4,3.9)", 691.698_36, 1e-12);
}

#[test]
fn db_errors() {
    fails("DB(1000000,100000,6,8,7)", ErrorValue::Num);
    fails("DB(1000000,100000,6,0)", ErrorValue::Num);
    fails("DB(1000000,100000,0,1)", ErrorValue::Num);
    fails("DB(1000000,100000,6,1,13)", ErrorValue::Num);
    fails("DB(1000000,100000,6,1,0)", ErrorValue::Num);
    fails("DB(-1,0,6,1)", ErrorValue::Num);
    fails("DB(1000,-1,6,1)", ErrorValue::Num);
    fails("DB(1000,100,6)", ErrorValue::Value);
}

#[test]
fn ddb_microsoft_examples_and_oracles() {
    near("DDB(2400,300,10*365,1)", 1.32, 5e-3);
    near("DDB(2400,300,10*12,1,2)", 40.0, 5e-3);
    near("DDB(2400,300,10,1,2)", 480.0, 5e-3);
    near("DDB(2400,300,10,2,1.5)", 306.0, 5e-3);
    near("DDB(2400,300,10,10)", 22.12, 5e-3);
    // mpmath 30 digits
    close("DDB(2400,300,3650,1)", 1.315_068_493_150_684_9, 1e-12);
    close("DDB(2400,300,10,10)", 22.122_547_2, 1e-12);
    close("DDB(10000,1000,5,2.5)", 1_859.032_006_179_56, 1e-12);
    assert_eq!(eval("DDB(1000,100,1,1,3)"), number(900.0));
}

#[test]
fn ddb_errors() {
    fails("DDB(2400,300,10,11)", ErrorValue::Num);
    fails("DDB(2400,300,10,0)", ErrorValue::Num);
    fails("DDB(2400,300,10,1,0)", ErrorValue::Num);
    fails("DDB(2400,-1,10,1)", ErrorValue::Num);
    fails("DDB(-1,0,10,1)", ErrorValue::Num);
    fails("DDB(2400,300,0,1)", ErrorValue::Num);
    fails("DDB(2400,300,10)", ErrorValue::Value);
}

#[test]
fn vdb_microsoft_examples_and_oracles() {
    near("VDB(2400,300,10*365,0,1)", 1.32, 5e-3);
    near("VDB(2400,300,10*12,0,1)", 40.0, 5e-3);
    near("VDB(2400,300,10,0,1)", 480.0, 5e-3);
    near("VDB(2400,300,10*12,6,18)", 396.31, 5e-3);
    near("VDB(2400,300,10*12,6,18,1.5)", 311.81, 5e-3);
    near("VDB(2400,300,10,0,0.875,1.5)", 315.0, 5e-3);
    // mpmath 30 digits
    close("VDB(2400,300,3650,0,1)", 1.315_068_493_150_684_9, 1e-12);
    close("VDB(2400,300,120,6,18)", 396.306_053_264_750_86, 1e-12);
    close("VDB(2400,300,120,6,18,1.5)", 311.808_936_658_234_1, 1e-12);
    close("VDB(10000,1000,5,2.5,4.25,1.5)", 2_253.75, 1e-12);
}

#[test]
fn vdb_switches_to_straight_line_unless_told_not_to() {
    // mpmath 30 digits
    close("VDB(10000,0,5,3,5)", 2_160.0, 1e-12);
    close("VDB(10000,0,5,3,5,2,FALSE)", 2_160.0, 1e-12);
    close("VDB(10000,0,5,3,5,2,TRUE)", 1_382.4, 1e-12);
    close("VDB(10000,0,5,0,5,2,TRUE)", 9_222.4, 1e-12);
    close("VDB(10000,0,5,3.5,4.5)", 1_080.0, 1e-12);
    assert_eq!(eval("VDB(10000,0,5,2,2)"), number(0.0));
}

#[test]
fn vdb_errors() {
    fails("VDB(2400,300,10,5,4)", ErrorValue::Num);
    fails("VDB(2400,300,10,0,11)", ErrorValue::Num);
    fails("VDB(2400,300,10,-1,1)", ErrorValue::Num);
    fails("VDB(2400,300,10,0,1,0)", ErrorValue::Num);
    fails("VDB(-1,0,10,0,1)", ErrorValue::Num);
    fails("VDB(2400,-1,10,0,1)", ErrorValue::Num);
    fails("VDB(2400,300,10,0)", ErrorValue::Value);
}

#[test]
fn effect_and_nominal_microsoft_examples_and_oracles() {
    near("EFFECT(0.0525,4)", 0.053_542_667, 5e-10);
    near("NOMINAL(0.053543,4)", 0.052_500_32, 5e-9);
    // mpmath 30 digits
    close("EFFECT(0.0525,4)", 0.053_542_667_370_758_06, 1e-12);
    close("EFFECT(0.12,12.9)", 0.126_825_030_131_969_72, 1e-12);
    close("NOMINAL(0.053543,4)", 0.052_500_319_868_355_86, 1e-12);
    close("NOMINAL(0.1,365)", 0.095_322_624_764_751_44, 1e-12);
    for name in ["EFFECT", "NOMINAL"] {
        fails(&format!("{name}(0,4)"), ErrorValue::Num);
        fails(&format!("{name}(-0.1,4)"), ErrorValue::Num);
        fails(&format!("{name}(0.1,0.9)"), ErrorValue::Num);
        fails(&format!("{name}(0.1)"), ErrorValue::Value);
        fails(&format!("{name}(\"x\",4)"), ErrorValue::Value);
    }
}

#[test]
fn pduration_and_rri_microsoft_examples_and_oracles() {
    near("PDURATION(2.5%,2000,2200)", 3.86, 5e-3);
    near("PDURATION(0.025/12,1000,1200)", 87.6, 5e-2);
    near("RRI(96,10000,11000)", 0.000_993_3, 5e-8);
    // mpmath 30 digits
    close("PDURATION(0.025,2000,2200)", 3.859_866_162_622_645_4, 1e-12);
    close(
        "PDURATION(0.025/12,1000,1200)",
        87.605_476_419_375_63,
        1e-12,
    );
    close("RRI(96,10000,11000)", 0.000_993_307_376_291_394_9, 1e-12);
    close("RRI(10,500,250)", -0.066_967_008_463_192_59, 1e-12);
    for args in ["0,2000,2200", "0.1,0,2200", "0.1,2000,-1", "-0.1,2000,2200"] {
        fails(&format!("PDURATION({args})"), ErrorValue::Num);
    }
    fails("RRI(0,10000,11000)", ErrorValue::Num);
    fails("RRI(10,0,11000)", ErrorValue::Num);
    fails("RRI(10,-100,100)", ErrorValue::Num);
    fails("RRI(10,100)", ErrorValue::Value);
}

#[test]
fn fvschedule_microsoft_example_and_oracle() {
    // mpmath 30 digits
    close("FVSCHEDULE(1,{0.09,0.11,0.1})", 1.330_89, 1e-12);
    let inputs = [
        ("A1".to_string(), number(0.05)),
        ("A2".to_string(), number(-0.02)),
        ("A4".to_string(), number(0.03)),
    ];
    close_in(&inputs, "FVSCHEDULE(1000,A1:A4)", 1_059.87, 1e-12);
    let with_text = [inputs.to_vec(), vec![("A3".to_string(), text("x"))]].concat();
    fails_in(&with_text, "FVSCHEDULE(1000,A1:A4)", ErrorValue::Value);
    let with_error = [
        inputs.to_vec(),
        vec![(
            "A3".to_string(),
            CellValue::Error {
                value: ErrorValue::Div0,
            },
        )],
    ]
    .concat();
    fails_in(&with_error, "FVSCHEDULE(1000,A1:A4)", ErrorValue::Div0);
    fails("FVSCHEDULE(1000)", ErrorValue::Value);
}

#[test]
fn dollarde_and_dollarfr_microsoft_examples() {
    close("DOLLARDE(1.02,16)", 1.125, 1e-12);
    close("DOLLARDE(1.1,32)", 1.3125, 1e-12);
    close("DOLLARFR(1.125,16)", 1.02, 1e-12);
    close("DOLLARFR(1.125,32)", 1.04, 1e-12);
    // mpmath 30 digits
    close("DOLLARDE(-3.5,8.9)", -3.625, 1e-12);
    close("DOLLARDE(2.123,1000)", 2.123, 1e-12);
    close("DOLLARFR(-7.25,4)", -7.1, 1e-12);
    for name in ["DOLLARDE", "DOLLARFR"] {
        fails(&format!("{name}(1.02,-1)"), ErrorValue::Num);
        fails(&format!("{name}(1.02,0.5)"), ErrorValue::Div0);
        fails(&format!("{name}(1.02,0)"), ErrorValue::Div0);
        fails(&format!("{name}(1.02)"), ErrorValue::Value);
    }
}

#[test]
fn rate_needs_a_positive_term() {
    fails("RATE(-5,0,-1000,1500)", ErrorValue::Num);
    fails("RATE(-1,0,-100,200)", ErrorValue::Num);
}

/// MIRR discounts by `(1 + rate)^t`, so a rate of -100% divides by zero.
#[test]
fn mirr_at_minus_one_hundred_percent_divides_by_zero() {
    fails("MIRR({100,-200},-1,0.1)", ErrorValue::Div0);
    fails("MIRR({-100,120},-1,0.1)", ErrorValue::Div0);
    fails("MIRR({-100,120},0.1,-1)", ErrorValue::Div0);
}

/// periods past 2^53 no longer step in floating point; the sum still ends.
#[test]
fn cumulative_periods_beyond_float_steps_terminate() {
    fails("CUMIPMT(0.1,1E16,1000,1E16,1E16,0)", ErrorValue::Num);
    fails("CUMPRINC(0.1,1E16,1000,1E16,1E16,0)", ErrorValue::Num);
}

/// the stub year after `life` exists only when the first year is short.
#[test]
fn db_has_no_stub_year_after_a_full_first_year() {
    fails("DB(1000000,100000,6,7)", ErrorValue::Num);
    fails("DB(1000000,100000,6,7,12)", ErrorValue::Num);
    close("DB(1000000,100000,6,7,7)", 15_845.098_473_848_073, 1e-12);
}

/// a salvage at or above cost, a zero cost included, depreciates nothing.
#[test]
fn db_without_anything_to_depreciate_is_zero() {
    assert_eq!(eval("DB(0,0,5,1)"), number(0.0));
    assert_eq!(eval("DB(100,200,5,1)"), number(0.0));
    assert_eq!(eval("DB(100,100,5,3)"), number(0.0));
}

#[test]
fn rri_rejects_non_positive_present_values_and_negative_futures() {
    fails("RRI(10,-100,-200)", ErrorValue::Num);
    fails("RRI(1,-100,100)", ErrorValue::Num);
    fails("RRI(1,100,-100)", ErrorValue::Num);
    assert_eq!(eval("RRI(2,100,0)"), number(-1.0));
}
