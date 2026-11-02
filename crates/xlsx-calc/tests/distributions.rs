//! GAMMA, beta, error, normal, lognormal, chi-square, t and f functions
//! against Microsoft's worked examples and independent oracle values.
//! every non-trivial expected value is an mpmath 1.4.1 result at 30
//! significant digits (40 working digits) quoted as the nearest f64; the
//! microsoft examples are checked at that precision rather than at the
//! rounded figures the reference prints, and the hypothesis test p-values
//! are mpmath evaluations of the textbook statistics over the surviving
//! pairs.

use xlsx_calc::{EvalContext, Value, evaluate, evaluate_array, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

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

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn eval_in(wb: &Workbook, src: &str) -> CellValue {
    let expr = parse_formula(src).expect("parse");
    evaluate(&expr, &EvalContext::new(wb, SheetId(0)))
}

fn approx(src: &str, want: f64) {
    approx_tol(src, want, 1e-12);
}

fn approx_inv(src: &str, want: f64) {
    approx_tol(src, want, 1e-9);
}

fn approx_tol(src: &str, want: f64, tol: f64) {
    let wb = workbook(&[]);
    match eval_in(&wb, src) {
        CellValue::Number { value } => {
            let scale = want.abs().max(1.0);
            assert!(
                (value - want).abs() <= tol * scale,
                "formula {src:?}: {value} != {want}"
            );
        }
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

fn check(cases: &[(&str, CellValue)]) {
    let wb = workbook(&[]);
    for (src, want) in cases {
        assert_eq!(eval_in(&wb, src), *want, "formula {src:?}");
    }
}

fn data_workbook() -> Workbook {
    workbook(&[
        ("A1", n(3.0)),
        ("A2", n(4.0)),
        ("A3", n(5.0)),
        ("A4", n(8.0)),
        ("A5", n(9.0)),
        ("A6", n(1.0)),
        ("A7", n(2.0)),
        ("A8", n(4.0)),
        ("A9", n(5.0)),
        ("B1", n(6.0)),
        ("B2", n(19.0)),
        ("B3", n(3.0)),
        ("B4", n(2.0)),
        ("B5", n(14.0)),
        ("B6", n(5.0)),
        ("B7", n(5.0)),
        ("B8", n(17.0)),
        ("B9", n(1.0)),
        ("D1", n(21.0)),
        ("D2", n(31.0)),
        ("D3", n(48.0)),
        ("E1", n(25.0)),
        ("E2", n(25.0)),
        ("E3", n(50.0)),
        ("G1", n(6.0)),
        ("G2", n(7.0)),
        ("G3", n(3.0)),
        ("G4", n(15.0)),
        ("G5", n(21.0)),
        ("H1", n(20.0)),
        ("H2", n(28.0)),
        ("H3", n(31.0)),
        ("H4", n(38.0)),
        ("H5", n(40.0)),
    ])
}

fn approx_in(wb: &Workbook, src: &str, want: f64, tol: f64) {
    match eval_in(wb, src) {
        CellValue::Number { value } => {
            let scale = want.abs().max(1.0);
            assert!(
                (value - want).abs() <= tol * scale,
                "formula {src:?}: {value} != {want}"
            );
        }
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

#[test]
fn gamma_family() {
    approx("GAMMA(5)", 24.0);
    approx("GAMMA(0.5)", 1.772453850905516);
    approx("GAMMA(-1.5)", 2.363271801207355);
    // mpmath 30 digits
    approx("GAMMA(2.5)", 1.329340388179137);
    // mpmath 30 digits
    approx("GAMMA(-2.5)", -0.9453087204829419);
    approx("GAMMALN(4)", 1.791759469228055);
    // mpmath 30 digits
    approx("GAMMALN(12.5)", 18.734347511936445);
    approx("GAMMALN.PRECISE(4)", 1.791759469228055);
    approx("IFERROR(GAMMA(5),0)", 24.0);
    approx("_xlfn.GAMMA(5)", 24.0);
    check(&[
        ("GAMMA(0)", e(ErrorValue::Num)),
        ("GAMMA(-2)", e(ErrorValue::Num)),
        ("GAMMA(\"x\")", e(ErrorValue::Value)),
        ("GAMMALN(0)", e(ErrorValue::Num)),
        ("GAMMALN(-1)", e(ErrorValue::Num)),
        ("GAMMALN.PRECISE(0)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn gamma_dist_and_inv() {
    // microsoft documentation example
    approx("GAMMA.DIST(10.00001131,9,2,FALSE)", 0.032639130418294);
    approx("GAMMA.DIST(10.00001131,9,2,TRUE)", 0.06809400386978733);
    approx("GAMMADIST(10.00001131,9,2,FALSE)", 0.032639130418294);
    approx("GAMMADIST(10.00001131,9,2,TRUE)", 0.06809400386978733);
    // mpmath 30 digits
    approx("GAMMA.DIST(5,2,3,TRUE)", 0.4963317257665018);
    // mpmath 30 digits
    approx("GAMMA.DIST(5,2,3,FALSE)", 0.10493089046531212);
    approx_inv("GAMMA.INV(0.068094,9,2)", 10.000011191437178);
    approx_inv("GAMMAINV(0.068094,9,2)", 10.000011191437178);
    // mpmath 30 digits
    approx_inv("GAMMA.INV(0.4,2,3)", 4.12926402618866);
    check(&[
        ("GAMMA.DIST(-1,9,2,TRUE)", e(ErrorValue::Num)),
        ("GAMMA.DIST(10,0,2,TRUE)", e(ErrorValue::Num)),
        ("GAMMA.DIST(10,9,0,TRUE)", e(ErrorValue::Num)),
        ("GAMMADIST(-1,9,2,TRUE)", e(ErrorValue::Num)),
        ("GAMMA.INV(-0.1,9,2)", e(ErrorValue::Num)),
        ("GAMMA.INV(1.1,9,2)", e(ErrorValue::Num)),
        ("GAMMA.INV(0.5,0,2)", e(ErrorValue::Num)),
        ("GAMMA.INV(0.5,9,0)", e(ErrorValue::Num)),
        ("GAMMAINV(1.1,9,2)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn beta_dist_and_inv() {
    // microsoft documentation example
    approx("BETA.DIST(2,8,10,TRUE,1,3)", 0.6854705810546875);
    approx("BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646484375);
    approx("BETADIST(2,8,10,1,3)", 0.6854705810546875);
    // mpmath 30 digits
    approx("BETA.DIST(0.3,2,5,TRUE)", 0.5798249999999999);
    // mpmath 30 digits
    approx("BETA.DIST(0.3,2,5,FALSE)", 2.1609000000000003);
    approx_inv("BETA.INV(0.685470581,8,10,1,3)", 1.9999999999631428);
    approx_inv("BETAINV(0.685470581,8,10,1,3)", 1.9999999999631428);
    // mpmath 30 digits
    approx_inv("BETA.INV(0.25,2,5)", 0.16116291679032652);
    check(&[
        ("BETA.DIST(2,0,10,TRUE,1,3)", e(ErrorValue::Num)),
        ("BETA.DIST(2,8,0,TRUE,1,3)", e(ErrorValue::Num)),
        ("BETA.DIST(0.5,8,10,TRUE,1,3)", e(ErrorValue::Num)),
        ("BETA.DIST(4,8,10,TRUE,1,3)", e(ErrorValue::Num)),
        ("BETADIST(0.5,8,10,1,3)", e(ErrorValue::Num)),
        ("BETA.INV(-0.1,8,10)", e(ErrorValue::Num)),
        ("BETA.INV(1.1,8,10)", e(ErrorValue::Num)),
        ("BETA.INV(0.5,0,10)", e(ErrorValue::Num)),
        ("BETAINV(1.1,8,10)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn erf_family() {
    // microsoft documentation example
    approx("ERF(0.745)", 0.7079289200957377);
    approx("ERF.PRECISE(0.745)", 0.7079289200957377);
    approx("ERFC(1)", 0.15729920705028513);
    approx("ERFC.PRECISE(1)", 0.15729920705028513);
    // mpmath 30 digits
    approx("ERF(0.2,1)", 0.6199982037392364);
    // mpmath 30 digits
    approx("ERFC(2.5)", 0.0004069520174449589);
    approx("ERF(-0.745)", -0.7079289200957377);
    check(&[("ERF(\"x\")", e(ErrorValue::Value))]);
}

#[test]
fn norm_family() {
    // microsoft documentation examples
    approx("NORM.S.DIST(1.333333,TRUE)", 0.9087887256040951);
    approx("NORMSDIST(1.333333)", 0.9087887256040951);
    approx("NORM.S.DIST(1.333333,FALSE)", 0.16401014756936722);
    approx("PHI(0.75)", 0.30113743215480443);
    approx("GAUSS(2)", 0.4772498680518208);
    // mpmath 30 digits
    approx("GAUSS(-1)", -0.3413447460685429);
    approx_inv("NORM.S.INV(0.908789)", 1.3333346730441074);
    approx_inv("NORMSINV(0.908789)", 1.3333346730441074);
    // mpmath 30 digits
    approx_inv("NORM.S.INV(0.125)", -1.150349380376008);
    approx_inv("NORM.INV(0.908789,40,1.5)", 42.00000200956616);
    approx_inv("NORMINV(0.908789,40,1.5)", 42.00000200956616);
    check(&[
        ("NORM.S.INV(0)", e(ErrorValue::Num)),
        ("NORM.S.INV(1)", e(ErrorValue::Num)),
        ("NORMSINV(0)", e(ErrorValue::Num)),
        ("NORM.INV(0,40,1.5)", e(ErrorValue::Num)),
        ("NORM.INV(0.5,40,0)", e(ErrorValue::Num)),
        ("NORM.INV(0.5,40,-1)", e(ErrorValue::Num)),
        ("NORMINV(1,40,1.5)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn lognorm_family() {
    // microsoft documentation examples
    approx("LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.039083555706800464);
    approx("LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.017617596681819222);
    approx("LOGNORMDIST(4,3.5,1.2)", 0.039083555706800464);
    // mpmath 30 digits
    approx("LOGNORM.DIST(2,1,0.5,TRUE)", 0.26970493073490953);
    approx_inv("LOGNORM.INV(0.039084,3.5,1.2)", 4.000025218680635);
    approx_inv("LOGINV(0.039084,3.5,1.2)", 4.000025218680635);
    check(&[
        ("LOGNORM.DIST(0,3.5,1.2,TRUE)", e(ErrorValue::Num)),
        ("LOGNORM.DIST(-1,3.5,1.2,TRUE)", e(ErrorValue::Num)),
        ("LOGNORM.DIST(4,3.5,0,TRUE)", e(ErrorValue::Num)),
        ("LOGNORMDIST(0,3.5,1.2)", e(ErrorValue::Num)),
        ("LOGNORM.INV(0,3.5,1.2)", e(ErrorValue::Num)),
        ("LOGNORM.INV(1,3.5,1.2)", e(ErrorValue::Num)),
        ("LOGNORM.INV(0.5,3.5,0)", e(ErrorValue::Num)),
        ("LOGINV(0,3.5,1.2)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn chisq_family() {
    // microsoft documentation examples
    approx("CHISQ.DIST(0.5,1,TRUE)", 0.5204998778130465);
    approx("CHISQ.DIST(0.5,1,FALSE)", 0.4393912894677224);
    approx("CHISQ.DIST.RT(18.307,10)", 0.05000058909139812);
    approx("CHIDIST(18.307,10)", 0.05000058909139812);
    // mpmath 30 digits
    approx("CHISQ.DIST(8,5,TRUE)", 0.8437643724222776);
    approx_inv("CHISQ.INV(0.05,10)", 3.94029913611906);
    approx_inv("CHISQ.INV.RT(0.05,10)", 18.307038053275146);
    approx_inv("CHIINV(0.05,10)", 18.307038053275146);
    check(&[
        ("CHISQ.DIST(-1,1,TRUE)", e(ErrorValue::Num)),
        ("CHISQ.DIST(0.5,0,TRUE)", e(ErrorValue::Num)),
        ("CHISQ.DIST.RT(-1,10)", e(ErrorValue::Num)),
        ("CHIDIST(-1,10)", e(ErrorValue::Num)),
        ("CHISQ.INV(-0.1,10)", e(ErrorValue::Num)),
        ("CHISQ.INV(1.1,10)", e(ErrorValue::Num)),
        ("CHISQ.INV(0.5,0)", e(ErrorValue::Num)),
        ("CHISQ.INV.RT(-0.1,10)", e(ErrorValue::Num)),
        ("CHIINV(1.1,10)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn t_family() {
    // microsoft documentation examples
    approx("T.DIST(1.959999998,60,TRUE)", 0.9726775350120396);
    approx("T.DIST(1.959999998,60,FALSE)", 0.059847906435580384);
    approx("T.DIST.RT(1.959999998,60)", 0.027322464987960433);
    approx("T.DIST.2T(1.959999998,60)", 0.05464492997592087);
    approx("TDIST(1.959999998,60,2)", 0.05464492997592087);
    approx("TDIST(1.959999998,60,1)", 0.027322464987960433);
    // mpmath 30 digits
    approx("T.DIST(-1.5,10,TRUE)", 0.0822536632227201);
    approx_inv("T.INV(0.75,2)", 0.816496580927726);
    approx_inv("T.INV.2T(0.05,60)", 2.0002978220142604);
    approx_inv("TINV(0.05,60)", 2.0002978220142604);
    check(&[
        ("T.DIST(1,0,TRUE)", e(ErrorValue::Num)),
        ("T.DIST.2T(-1,60)", e(ErrorValue::Num)),
        ("TDIST(-1,60,2)", e(ErrorValue::Num)),
        ("TDIST(1,60,3)", e(ErrorValue::Num)),
        ("TDIST(1,0,2)", e(ErrorValue::Num)),
        ("T.INV(0,10)", e(ErrorValue::Num)),
        ("T.INV(1,10)", e(ErrorValue::Num)),
        ("T.INV(0.5,0)", e(ErrorValue::Num)),
        ("T.INV.2T(0,60)", e(ErrorValue::Num)),
        ("T.INV.2T(1.1,60)", e(ErrorValue::Num)),
        ("TINV(0,60)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn f_family() {
    // microsoft documentation examples
    approx("F.DIST(15.2068649,6,4,TRUE)", 0.9900000000475354);
    approx("F.DIST(15.2068649,6,4,FALSE)", 0.0012237995897796605);
    approx("F.DIST.RT(15.2068649,6,4)", 0.009999999952464601);
    approx("FDIST(15.2068649,6,4)", 0.009999999952464601);
    // mpmath 30 digits
    approx_inv("F.INV(0.5,6,4)", 1.0616688782738317);
    approx_inv("F.INV.RT(0.05,6,4)", 6.163132282688634);
    approx_inv("FINV(0.05,6,4)", 6.163132282688634);
    check(&[
        ("F.DIST(-1,6,4,TRUE)", e(ErrorValue::Num)),
        ("F.DIST(1,0,4,TRUE)", e(ErrorValue::Num)),
        ("F.DIST(1,6,0,TRUE)", e(ErrorValue::Num)),
        ("FDIST(-1,6,4)", e(ErrorValue::Num)),
        ("F.INV(-0.1,6,4)", e(ErrorValue::Num)),
        ("F.INV(1.1,6,4)", e(ErrorValue::Num)),
        ("F.INV(0.5,0,4)", e(ErrorValue::Num)),
        ("F.INV.RT(-0.1,6,4)", e(ErrorValue::Num)),
        ("FINV(1.1,6,4)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn confidence_intervals() {
    // microsoft documentation examples
    approx_inv("CONFIDENCE.NORM(0.05,2.5,50)", 0.692951912174839);
    approx_inv("CONFIDENCE(0.05,2.5,50)", 0.692951912174839);
    approx_inv("CONFIDENCE.T(0.05,1,50)", 0.2841968554957299);
    check(&[
        ("CONFIDENCE.NORM(0,2.5,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.NORM(1,2.5,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.NORM(0.05,0,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.NORM(0.05,2.5,0)", e(ErrorValue::Num)),
        ("CONFIDENCE(0,2.5,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.T(0,1,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.T(1,1,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.T(0.05,0,50)", e(ErrorValue::Num)),
        ("CONFIDENCE.T(0.05,1,0)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn hypothesis_tests() {
    let wb = data_workbook();
    // mpmath 30 digits
    approx_in(&wb, "T.TEST(A1:A9,B1:B9,2,2)", 0.1750801005561359, 1e-9);
    approx_in(&wb, "TTEST(A1:A9,B1:B9,2,2)", 0.1750801005561359, 1e-9);
    approx_in(&wb, "T.TEST(A1:A9,B1:B9,2,3)", 0.1854508930877836, 1e-9);
    approx_in(&wb, "T.TEST(A1:A5,B1:B5,2,1)", 0.4472750443620635, 1e-9);
    approx_in(&wb, "CHISQ.TEST(D1:D3,E1:E3)", 0.33959552564493917, 1e-9);
    approx_in(&wb, "CHITEST(D1:D3,E1:E3)", 0.33959552564493917, 1e-9);
    approx_in(&wb, "F.TEST(G1:G5,H1:H5)", 0.8748742045675306, 1e-9);
    approx_in(&wb, "FTEST(G1:G5,H1:H5)", 0.8748742045675306, 1e-9);
    assert_eq!(eval_in(&wb, "T.TEST(A1:A9,B1:B9,2,4)"), e(ErrorValue::Num));
    assert_eq!(eval_in(&wb, "T.TEST(A1:A9,B1:B9,3,2)"), e(ErrorValue::Num));
    assert_eq!(eval_in(&wb, "T.TEST(A1:A5,B1:B9,2,1)"), e(ErrorValue::NA));
    assert_eq!(eval_in(&wb, "CHISQ.TEST(D1:D3,E1:E2)"), e(ErrorValue::NA));
}

#[test]
fn aliases_match_modern_names() {
    let wb = workbook(&[]);
    for (modern, legacy) in [
        ("GAMMA.DIST(5,2,3,TRUE)", "GAMMADIST(5,2,3,TRUE)"),
        ("GAMMA.INV(0.4,2,3)", "GAMMAINV(0.4,2,3)"),
        ("BETA.INV(0.25,2,5)", "BETAINV(0.25,2,5)"),
        ("NORM.INV(0.3,40,1.5)", "NORMINV(0.3,40,1.5)"),
        ("NORM.S.INV(0.3)", "NORMSINV(0.3)"),
        ("LOGNORM.INV(0.3,1,0.5)", "LOGINV(0.3,1,0.5)"),
        ("CHISQ.DIST.RT(8,5)", "CHIDIST(8,5)"),
        ("CHISQ.INV.RT(0.3,5)", "CHIINV(0.3,5)"),
        ("T.INV.2T(0.3,10)", "TINV(0.3,10)"),
        ("F.DIST.RT(2,6,4)", "FDIST(2,6,4)"),
        ("F.INV.RT(0.3,6,4)", "FINV(0.3,6,4)"),
        ("CONFIDENCE.NORM(0.05,2.5,50)", "CONFIDENCE(0.05,2.5,50)"),
        ("GAMMALN(4)", "GAMMALN.PRECISE(4)"),
        ("ERF(0.5)", "ERF.PRECISE(0.5)"),
        ("ERFC(0.5)", "ERFC.PRECISE(0.5)"),
    ] {
        let value = eval_in(&wb, modern);
        assert!(matches!(value, CellValue::Number { .. }), "{modern}");
        assert_eq!(value, eval_in(&wb, legacy), "{modern} vs {legacy}");
    }
    let wb = data_workbook();
    for (modern, legacy) in [
        ("CHISQ.TEST(D1:D3,E1:E3)", "CHITEST(D1:D3,E1:E3)"),
        ("T.TEST(A1:A9,B1:B9,2,2)", "TTEST(A1:A9,B1:B9,2,2)"),
        ("F.TEST(G1:G5,H1:H5)", "FTEST(G1:G5,H1:H5)"),
    ] {
        let value = eval_in(&wb, modern);
        assert!(matches!(value, CellValue::Number { .. }), "{modern}");
        assert_eq!(value, eval_in(&wb, legacy), "{modern} vs {legacy}");
    }
}

fn t(v: &str) -> CellValue {
    CellValue::Text { value: v.into() }
}

/// J/K and L/M hold blanks and text at different positions, so pairing the
/// numbers in order would mismatch them; P/Q skip a text and a blank cell.
fn gappy_workbook() -> Workbook {
    workbook(&[
        ("J1", n(3.0)),
        ("J3", n(5.0)),
        ("J4", t("x")),
        ("J5", n(9.0)),
        ("J6", n(2.0)),
        ("J7", n(7.0)),
        ("K1", n(6.0)),
        ("K2", n(19.0)),
        ("K3", n(4.0)),
        ("K4", n(2.0)),
        ("K6", n(5.0)),
        ("K7", n(11.0)),
        ("L1", n(3.0)),
        ("L3", n(5.0)),
        ("L4", n(9.0)),
        ("L5", n(2.0)),
        ("L6", n(8.0)),
        ("M1", n(6.0)),
        ("M2", n(19.0)),
        ("M3", n(4.0)),
        ("M5", n(5.0)),
        ("M6", n(1.0)),
        ("P1", n(21.0)),
        ("P2", t("x")),
        ("P3", n(31.0)),
        ("P4", n(48.0)),
        ("Q1", n(25.0)),
        ("Q2", n(10.0)),
        ("Q3", n(25.0)),
        ("Q4", n(50.0)),
        ("Q5", n(7.0)),
    ])
}

#[test]
fn paired_t_test_pairs_by_position() {
    let wb = gappy_workbook();
    // mpmath 30 digits over the pairs (3,6) (5,4) (2,5) (7,11)
    approx_in(&wb, "T.TEST(J1:J7,K1:K7,2,1)", 0.1354165215079502, 1e-9);
    approx_in(&wb, "T.TEST(J1:J7,K1:K7,1,1)", 0.0677082607539751, 1e-9);
    approx_in(&wb, "TTEST(J1:J7,K1:K7,2,1)", 0.1354165215079502, 1e-9);
    // mpmath 30 digits over the pairs (3,6) (5,4) (2,5) (8,1)
    approx_in(&wb, "T.TEST(L1:L6,M1:M6,2,1)", 0.8459763034682742, 1e-9);
    // mpmath 30 digits over the pairs (3,6) (2,5) (7,11)
    approx_in(
        &wb,
        "T.TEST({3,\"x\",5,2,7},{6,4,\"y\",5,11},2,1)",
        0.00985245702332569,
        1e-9,
    );
    assert_eq!(eval_in(&wb, "T.TEST(J1:J7,K1:K6,2,1)"), e(ErrorValue::NA));
    assert_eq!(eval_in(&wb, "T.TEST({1,2,3},{4,5},2,1)"), e(ErrorValue::NA));
    assert_eq!(eval_in(&wb, "T.TEST(J1:J2,K1:K2,2,1)"), e(ErrorValue::Div0));
}

#[test]
fn chisq_test_skips_non_numeric_positions() {
    let wb = gappy_workbook();
    // mpmath 30 digits over (21,25) (31,25) (48,50) with df 4 from the shape
    approx_in(&wb, "CHISQ.TEST(P1:P5,Q1:Q5)", 0.7063586933414734, 1e-9);
    approx_in(&wb, "CHITEST(P1:P5,Q1:Q5)", 0.7063586933414734, 1e-9);
    assert_eq!(eval_in(&wb, "CHISQ.TEST(P1:P5,Q1:Q4)"), e(ErrorValue::NA));
    assert_eq!(
        eval_in(&wb, "CHISQ.TEST(P1:P5,{25,10,25,50,7})"),
        e(ErrorValue::NA)
    );
}

fn arrayed(wb: &Workbook, src: &str) -> (usize, usize, Vec<CellValue>) {
    let expr = parse_formula(src).expect("parse");
    match evaluate_array(&expr, &EvalContext::new(wb, SheetId(0))) {
        Value::Array(array) => (array.rows(), array.cols(), array.into_values()),
        other => (1, 1, vec![other.into_scalar()]),
    }
}

fn close(src: &str, got: &CellValue, want: &CellValue, tol: f64) {
    match (got, want) {
        (CellValue::Number { value }, CellValue::Number { value: want }) => {
            let scale = want.abs().max(1.0);
            assert!(
                (value - want).abs() <= tol * scale,
                "formula {src:?}: {value} != {want}"
            );
        }
        _ => assert_eq!(got, want, "formula {src:?}"),
    }
}

fn check_array(wb: &Workbook, src: &str, rows: usize, cols: usize, want: &[CellValue], tol: f64) {
    let (r, c, values) = arrayed(wb, src);
    assert_eq!((r, c), (rows, cols), "formula {src:?}");
    assert_eq!(values.len(), want.len(), "formula {src:?}");
    for (got, want) in values.iter().zip(want) {
        close(src, got, want, tol);
    }
}

#[test]
fn scalar_functions_lift_over_arrays() {
    let wb = data_workbook();
    // mpmath 30 digits
    check_array(
        &wb,
        "NORM.S.DIST({-1,0,1},TRUE)",
        1,
        3,
        &[n(0.15865525393145705), n(0.5), n(0.8413447460685429)],
        1e-12,
    );
    check_array(&wb, "GAMMA(A1:A3)", 3, 1, &[n(2.0), n(6.0), n(24.0)], 1e-12);
    // mpmath 30 digits
    check_array(
        &wb,
        "T.DIST(1.5,{1,2,10},TRUE)",
        1,
        3,
        &[
            n(0.8128329581890011),
            n(0.8638034375544995),
            n(0.9177463367772799),
        ],
        1e-12,
    );
    // mpmath 30 digits
    check_array(
        &wb,
        "CHIDIST({1;2;3},5)",
        3,
        1,
        &[
            n(0.9625657732472964),
            n(0.8491450360846097),
            n(0.6999858358786275),
        ],
        1e-12,
    );
    // mpmath 30 digits
    check_array(
        &wb,
        "BETA.DIST({0.2;0.5},{2,3},5,TRUE)",
        2,
        2,
        &[
            n(0.34464),
            n(0.14803200000000002),
            n(0.890625),
            n(0.7734375),
        ],
        1e-12,
    );
    check_array(
        &wb,
        "GAMMA({5,0,-2,0.5})",
        1,
        4,
        &[
            n(24.0),
            e(ErrorValue::Num),
            e(ErrorValue::Num),
            n(1.772453850905516),
        ],
        1e-12,
    );
    check_array(
        &wb,
        "NORM.S.INV({0.5,0,\"x\"})",
        1,
        3,
        &[n(0.0), e(ErrorValue::Num), e(ErrorValue::Value)],
        1e-9,
    );
    check_array(
        &wb,
        "ERFC.PRECISE({1,-1})",
        1,
        2,
        &[n(0.15729920705028513), n(1.8427007929497148)],
        1e-12,
    );
}

#[test]
fn hypothesis_tests_keep_ranges_whole_in_arrays() {
    let wb = data_workbook();
    // mpmath 30 digits
    check_array(
        &wb,
        "T.TEST(A1:A9,B1:B9,{1,2},2)",
        1,
        2,
        &[n(0.08754005027806795), n(0.1750801005561359)],
        1e-9,
    );
    check_array(
        &wb,
        "T.TEST(A1:A9,B1:B9,{2,4},2)",
        1,
        2,
        &[n(0.1750801005561359), e(ErrorValue::Num)],
        1e-9,
    );
    check_array(
        &wb,
        "CHISQ.TEST(D1:D3,E1:E3)",
        1,
        1,
        &[n(0.33959552564493917)],
        1e-9,
    );
    check_array(
        &wb,
        "F.TEST(G1:G5,H1:H5)",
        1,
        1,
        &[n(0.8748742045675306)],
        1e-9,
    );
}

/// relative error against `want`, for values far below 1 where the absolute
/// checks above would pass anything.
fn rel(src: &str, want: f64, tol: f64) {
    let wb = workbook(&[]);
    match eval_in(&wb, src) {
        CellValue::Number { value } => assert!(
            ((value - want) / want).abs() <= tol,
            "formula {src:?}: {value} != {want}"
        ),
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

#[test]
fn beta_density_at_the_endpoints() {
    check(&[
        ("BETA.DIST(0,1,5,FALSE)", n(5.0)),
        ("BETA.DIST(1,5,1,FALSE)", n(5.0)),
        ("BETA.DIST(2,1,5,FALSE,2,4)", n(2.5)),
        ("BETA.DIST(0,2,5,FALSE)", n(0.0)),
        ("BETA.DIST(0,0.5,5,FALSE)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn small_quantiles_converge_relatively() {
    // mpmath 30 digits
    rel("CHISQ.INV(1E-8,1)", 1.5707963267948967e-16, 1e-9);
    rel("BETA.INV(0.001,0.2,1)", 1.000000000000002e-15, 1e-9);
    rel("GAMMA.INV(0.5,1,1E-14)", 6.931471805599453e-15, 1e-9);
    rel("GAMMA.INV(1E-300,2,1)", 1.4142135623730952e-150, 1e-9);
}

#[test]
fn right_tails_are_computed_directly() {
    // mpmath 30 digits
    rel("T.DIST.RT(9,100)", 7.680385257375208e-15, 1e-12);
    rel("T.DIST.RT(10,100)", 4.9508444922970696e-17, 1e-12);
    rel("T.DIST.2T(10,100)", 9.901688984594139e-17, 1e-12);
    rel("TDIST(10,100,2)", 9.901688984594139e-17, 1e-12);
    rel("TDIST(10,100,1)", 4.9508444922970696e-17, 1e-12);
    rel("T.DIST.RT(-3,5)", 0.9849503760512687, 1e-12);
    rel(
        "T.TEST({1,2,3,4,5},{31,32,33,34,36},2,2)",
        3.6807320012804017e-9,
        1e-9,
    );
    rel("F.DIST.RT(200,10,20)", 8.740075266067028e-18, 1e-12);
    rel("FDIST(200,10,20)", 8.740075266067028e-18, 1e-12);
    rel("F.DIST.RT(1E16,1,1)", 6.366197723675813e-9, 1e-12);
    rel(
        "F.TEST({1,2,3,4,5,6},{100,-300,50,900,10,-700})",
        5.980459539202626e-12,
        1e-9,
    );
    rel("CHISQ.DIST.RT(400,10)", 9.413291991183476e-80, 1e-12);
    rel("CHIDIST(400,10)", 9.413291991183476e-80, 1e-12);
    approx("F.DIST(1E307*10,2,2,TRUE)", 1.0);
}

#[test]
fn tail_inverses_take_the_tail_probability() {
    // mpmath 30 digits
    rel("CHISQ.INV.RT(1E-12,2)", 55.262042231857095, 1e-9);
    rel("CHISQ.INV.RT(1E-20,2)", 92.10340371976183, 1e-9);
    rel("T.INV.2T(1E-12,2)", 999999.99999925, 1e-9);
    rel("TINV(1E-12,2)", 999999.99999925, 1e-9);
    rel("F.INV.RT(1E-9,1,1)", 4.0528473456935104e17, 1e-9);
    rel("CONFIDENCE.NORM(1E-17,1,1)", 8.573944076720883, 1e-9);
    rel("CONFIDENCE.T(1E-10,1,5)", 221.33487793313847, 1e-9);
    rel("NORM.S.INV(0.999999999999999)", 7.941444487415978, 1e-9);
    rel("T.INV(0.999999,3)", 103.29946777942897, 1e-9);
}

#[test]
fn large_shape_and_degrees_of_freedom() {
    // mpmath 30 digits
    approx("CHISQ.DIST(1000000,1000000,TRUE)", 0.5001880631966055);
    approx("GAMMA.DIST(500000,500000,1,TRUE)", 0.5001880631966055);
    approx("CHISQ.DIST.RT(1000000,1000000)", 0.4998119368033945);
    approx("GAMMA.DIST(500700,500000,1,TRUE)", 0.8389029798454206);
    approx("CHISQ.DIST.RT(1001400,1000000)", 0.16109702015457938);
    rel(
        "CHISQ.DIST(996000,1000000,TRUE)",
        0.0023148054846039395,
        1e-12,
    );
    approx_inv("CHISQ.INV(0.5,1000000)", 999999.3333334123);
    approx("T.DIST(1,1E9,TRUE)", 0.8413447459475576);
    approx("T.DIST(1,1E16,TRUE)", 0.8413447460685429);
    approx("T.DIST(1,1E9,FALSE)", 0.24197072439815798);
}

#[test]
fn small_gauss_keeps_relative_precision() {
    // mpmath 30 digits
    rel("GAUSS(1E-10)", 3.989422804014327e-11, 1e-12);
}

#[test]
fn confidence_t_with_one_observation_divides_by_zero() {
    check(&[
        ("CONFIDENCE.T(0.05,1,1)", e(ErrorValue::Div0)),
        ("CONFIDENCE.T(0.05,1,0)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn gamma_dist_accepts_a_text_cumulative_flag() {
    // mpmath 30 digits
    approx("GAMMA.DIST(10.00001131,9,2,\"TRUE\")", 0.06809400386978733);
}

#[test]
fn erfc_accepts_negative_arguments() {
    // mpmath 30 digits; Microsoft's current ERFC and ERFC.PRECISE pages
    // place no sign restriction on x
    approx("ERFC(-1)", 1.8427007929497148);
    approx("ERFC.PRECISE(-1)", 1.8427007929497148);
    approx("ERFC(-0.3)", 1.3286267594591274);
    approx("ERFC(-1E200)", 2.0);
}

#[test]
fn unconverged_beta_fraction_is_an_error() {
    check(&[("BETA.DIST(0.5,5E11,5E11,TRUE)", e(ErrorValue::Num))]);
}

#[test]
fn inverses_at_huge_quantiles() {
    // mpmath 30 digits
    rel("F.INV.RT(1E-120,1,1)", 4.052847345693511e239, 1e-9);
    rel("T.INV(1E-300,1)", -3.183098861837907e299, 1e-9);
    check(&[
        ("F.INV.RT(1E-300,1,1)", e(ErrorValue::Num)),
        ("T.INV(2^-1074,1)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn huge_arguments_saturate() {
    check(&[
        ("ERF(1E200)", n(1.0)),
        ("ERF(-1E200)", n(-1.0)),
        ("ERFC(1E200)", n(0.0)),
        ("NORM.S.DIST(1E200,TRUE)", n(1.0)),
        ("NORM.S.DIST(-1E200,TRUE)", n(0.0)),
        ("GAMMA.DIST(1E307*10,1,1E-10,TRUE)", n(1.0)),
        ("GAMMA.DIST(1E307*10,1E307*10,1E-10,TRUE)", n(1.0)),
        ("T.DIST(1E154,5,TRUE)", n(1.0)),
        ("T.DIST.RT(1E154,5)", n(0.0)),
        ("F.DIST.RT(2^-1074,2,2)", n(1.0)),
        ("CHISQ.TEST({1E200,1},{1,1})", n(0.0)),
    ]);
}

#[test]
fn chisq_test_zero_expectation_divides_by_zero() {
    check(&[
        ("CHISQ.TEST({1,2},{0,2})", e(ErrorValue::Div0)),
        ("CHISQ.TEST({-1,2},{1,2})", e(ErrorValue::Num)),
        ("CHISQ.TEST({1,2},{-1,2})", e(ErrorValue::Num)),
    ]);
}

#[test]
fn tiny_arguments_keep_their_digits() {
    // mpmath 30 digits
    rel("ERF(1E-200)", 1.1283791670955125e-200, 1e-12);
    rel("ERF.PRECISE(1E-200)", 1.1283791670955125e-200, 1e-12);
    rel("GAUSS(1E-200)", 3.9894228040143265e-201, 1e-12);
    rel("GAMMALN(2^-1074)", 744.4400719213812, 1e-12);
    rel("GAMMALN(1E-200)", 460.51701859880916, 1e-12);
    rel("GAMMA.DIST(2^-1074,1,1,FALSE)", 1.0, 1e-12);
    rel("F.DIST(1E-200,2,2,TRUE)", 1e-200, 1e-12);
    approx("BETA.DIST(0.2,2^-1074,1,TRUE)", 1.0);
    check(&[
        ("ERF(0)", n(0.0)),
        ("GAMMA.DIST(2^-1074,2,1,FALSE)", n(5e-324)),
    ]);
}

#[test]
fn beta_bounds_spanning_the_whole_range() {
    approx("BETA.DIST(0,2,2,TRUE,-1E307*10,1E307*10)", 0.5);
    // mpmath 30 digits
    rel(
        "BETA.INV(0.25,2,2,-1E307*10,1E307*10)",
        -3.472963553338607e307,
        1e-9,
    );
}

#[test]
fn samples_at_extreme_magnitudes() {
    // mpmath 30 digits; the statistics are scale free
    approx(
        "T.TEST({1E200,2E200,3E200},{4E200,6E200,9E200},2,2)",
        0.05024811780069481,
    );
    approx(
        "T.TEST({1E-200,2E-200,3E-200},{4E-200,6E-200,9E-200},2,2)",
        0.050248117800694825,
    );
    approx(
        "F.TEST({1E200,2E200,4E200},{1E200,5E200,9E200})",
        0.2545454545454546,
    );
    approx(
        "T.TEST({1,-1,0.5}*(1E307*10),{-1,1,0}*(1E307*10),2,1)",
        0.8994962184740788,
    );
}

#[test]
fn documented_degrees_of_freedom_limits() {
    approx("CHISQ.DIST.RT(1,1E10)", 1.0);
    // mpmath 30 digits
    approx("T.DIST(1,1E300,TRUE)", 0.8413447460685429);
    check(&[
        ("CHISQ.DIST(1,1E10+1,TRUE)", e(ErrorValue::Num)),
        ("CHISQ.DIST.RT(1,2E10)", e(ErrorValue::Num)),
        ("CHIDIST(1,2E10)", e(ErrorValue::Num)),
        ("CHISQ.INV(0.5,2E10)", e(ErrorValue::Num)),
        ("CHISQ.INV.RT(0.5,2E10)", e(ErrorValue::Num)),
        ("CHIINV(0.5,2E10)", e(ErrorValue::Num)),
        ("F.DIST.RT(1,1E10+1,1)", e(ErrorValue::Num)),
        ("FDIST(1,1,2E10)", e(ErrorValue::Num)),
        ("F.INV.RT(0.5,2E10,2)", e(ErrorValue::Num)),
        ("FINV(0.5,2,1E10+1)", e(ErrorValue::Num)),
        ("F.DIST(1,2E10,1,TRUE)", e(ErrorValue::Num)),
        ("F.INV(0.5,2,2E10)", e(ErrorValue::Num)),
    ]);
}

/// Excel 365 for Mac result: F.INV.RT(0.5,1E10,1E10) is 1, so 10^10
/// degrees of freedom are still accepted by every f function.
#[test]
fn f_functions_accept_ten_billion_degrees_of_freedom() {
    approx_inv("F.INV.RT(0.5,1E10,1E10)", 1.0);
    approx_inv("FINV(0.5,1E10,1E10)", 1.0);
    approx_inv("F.INV(0.5,1E10,1E10)", 1.0);
    approx_inv("F.DIST(1,1E10,1E10,TRUE)", 0.5);
    approx_inv("F.DIST.RT(1,1E10,1E10)", 0.5);
    approx_inv("FDIST(1,1E10,1E10)", 0.5);
}

/// Excel 365 for Mac result: T.DIST.RT(1E200,5) is #NUM!. the t functions
/// square `x`, so they answer `#NUM!` once that overflows, from about
/// 1.34E+154; just below it the tails are still 0 and 1.
#[test]
fn t_functions_reject_x_whose_square_overflows() {
    check(&[
        ("T.DIST.RT(1E200,5)", e(ErrorValue::Num)),
        ("T.DIST(1E200,5,TRUE)", e(ErrorValue::Num)),
        ("T.DIST(-1E200,5,TRUE)", e(ErrorValue::Num)),
        ("T.DIST(1E200,5,FALSE)", e(ErrorValue::Num)),
        ("T.DIST.2T(1E200,5)", e(ErrorValue::Num)),
        ("TDIST(1E200,5,1)", e(ErrorValue::Num)),
        ("T.DIST.RT(1.35E154,5)", e(ErrorValue::Num)),
        ("T.DIST.RT(1.34E154,5)", n(0.0)),
        ("T.DIST(-1.34E154,5,TRUE)", n(0.0)),
        ("T.DIST.2T(1.34E154,5)", n(0.0)),
        ("TDIST(1.34E154,5,2)", n(0.0)),
    ]);
}
