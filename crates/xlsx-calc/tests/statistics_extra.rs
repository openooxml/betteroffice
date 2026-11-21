//! discrete distributions and the remaining statistics against Microsoft's
//! worked examples and independent oracle values. every non-trivial expected
//! value is an mpmath 1.4.1 result at 30 significant digits (40 working
//! digits) quoted as the nearest f64: the discrete masses and tails are exact
//! rational sums (or regularized incomplete beta and gamma functions) and the
//! statistics are the textbook formulas over the same data. the microsoft
//! examples are checked at that precision rather than at the rounded figures
//! the reference prints.

use xlsx_calc::{EvalContext, Value, evaluate, evaluate_array, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn t(s: &str) -> CellValue {
    CellValue::Text {
        value: s.to_string(),
    }
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

fn column(col: &str, values: &[f64]) -> Vec<(String, CellValue)> {
    values
        .iter()
        .enumerate()
        .map(|(i, v)| (format!("{col}{}", i + 1), n(*v)))
        .collect()
}

/// A: PERCENTRANK.INC data, B: PERCENTILE.EXC / PERCENTRANK.EXC data,
/// C: QUARTILE.EXC data, D/E: PROB, F: VARA tool strengths, G: Z.TEST data,
/// H/I: LOGEST months and units, J: MODE.MULT data, K: `-A` mixed cells.
fn data_workbook() -> Workbook {
    let mut cells: Vec<(String, CellValue)> = Vec::new();
    cells.extend(column(
        "A",
        &[13.0, 12.0, 11.0, 8.0, 4.0, 3.0, 2.0, 1.0, 1.0, 1.0],
    ));
    cells.extend(column("B", &[1.0, 2.0, 3.0, 6.0, 6.0, 6.0, 7.0, 8.0, 9.0]));
    cells.extend(column(
        "C",
        &[
            6.0, 7.0, 15.0, 36.0, 39.0, 40.0, 41.0, 42.0, 43.0, 47.0, 49.0,
        ],
    ));
    cells.extend(column("D", &[0.0, 1.0, 2.0, 3.0]));
    cells.extend(column("E", &[0.2, 0.3, 0.1, 0.4]));
    cells.extend(column(
        "F",
        &[
            1345.0, 1301.0, 1368.0, 1322.0, 1310.0, 1370.0, 1318.0, 1350.0, 1303.0, 1299.0,
        ],
    ));
    cells.extend(column(
        "G",
        &[3.0, 6.0, 7.0, 8.0, 6.0, 5.0, 4.0, 2.0, 1.0, 9.0],
    ));
    cells.extend(column(
        "H",
        &[11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0],
    ));
    cells.extend(column(
        "I",
        &[33100.0, 47300.0, 69000.0, 102000.0, 150000.0, 220000.0],
    ));
    cells.extend(column(
        "J",
        &[1.0, 2.0, 3.0, 4.0, 3.0, 2.0, 1.0, 2.0, 3.0, 5.0, 6.0, 1.0],
    ));
    cells.push(("K1".into(), n(1.0)));
    cells.push(("K2".into(), CellValue::Bool { value: true }));
    cells.push(("K3".into(), t("x")));
    cells.push(("K4".into(), n(4.0)));
    cells.push(("L1".into(), t("y")));
    let refs: Vec<(&str, CellValue)> = cells.iter().map(|(a, v)| (a.as_str(), v.clone())).collect();
    workbook(&refs)
}

fn eval_in(wb: &Workbook, src: &str) -> CellValue {
    let expr = parse_formula(src).expect("parse");
    evaluate(&expr, &EvalContext::new(wb, SheetId(0)))
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

/// relative 1e-12, the closed-form and series tolerance.
fn approx(src: &str, want: f64) {
    approx_in(&data_workbook(), src, want, 1e-12);
}

fn check(cases: &[(&str, CellValue)]) {
    let wb = data_workbook();
    for (src, want) in cases {
        assert_eq!(eval_in(&wb, src), *want, "formula {src:?}");
    }
}

fn arrayed(wb: &Workbook, src: &str) -> (usize, usize, Vec<CellValue>) {
    let expr = parse_formula(src).expect("parse");
    match evaluate_array(&expr, &EvalContext::new(wb, SheetId(0))) {
        Value::Array(array) => (array.rows(), array.cols(), array.into_values()),
        other => (1, 1, vec![other.into_scalar()]),
    }
}

fn check_array(src: &str, rows: usize, cols: usize, want: &[CellValue], tol: f64) {
    let (r, c, values) = arrayed(&data_workbook(), src);
    assert_eq!((r, c), (rows, cols), "formula {src:?}");
    assert_eq!(values.len(), want.len(), "formula {src:?}");
    for (got, want) in values.iter().zip(want) {
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
}

#[test]
fn binom_dist() {
    // microsoft documentation example
    approx("BINOM.DIST(6,10,0.5,FALSE)", 0.205078125);
    approx("BINOMDIST(6,10,0.5,FALSE)", 0.205078125);
    approx("_xlfn.BINOM.DIST(6,10,0.5,FALSE)", 0.205078125);
    // mpmath 30 digits
    approx("BINOM.DIST(6,10,0.5,TRUE)", 0.828125);
    // mpmath 30 digits
    approx("BINOM.DIST(30,100,0.37,TRUE)", 0.08771432738534025);
    // mpmath 30 digits
    approx("BINOM.DIST(30,100,0.37,FALSE)", 0.029364913634399917);
    // mpmath 30 digits, regularized incomplete beta
    approx("BINOM.DIST(4900,10000,0.5,TRUE)", 0.023292763852473693);
    // mpmath 30 digits
    approx("BINOM.DIST(510,1000,0.5,FALSE)", 0.020656351649512027);
    // mpmath 30 digits
    approx("BINOM.DIST(2,1000,0.001,TRUE)", 0.9197906571597989);
    approx("BINOM.DIST(6.9,10.7,0.5,FALSE)", 0.205078125);
    approx("BINOMDIST(30,100,0.37,TRUE)", 0.08771432738534025);
    approx("BINOM.DIST(0,5,0,FALSE)", 1.0);
    approx("BINOM.DIST(5,5,1,FALSE)", 1.0);
    check(&[
        ("BINOM.DIST(11,10,0.5,TRUE)", e(ErrorValue::Num)),
        ("BINOM.DIST(-1,10,0.5,TRUE)", e(ErrorValue::Num)),
        ("BINOM.DIST(5,10,1.5,TRUE)", e(ErrorValue::Num)),
        ("BINOM.DIST(5,10,-0.1,FALSE)", e(ErrorValue::Num)),
        ("BINOM.DIST(\"x\",10,0.5,TRUE)", e(ErrorValue::Value)),
        ("BINOM.DIST(5,10,0.5,\"x\")", e(ErrorValue::Value)),
        ("BINOMDIST(11,10,0.5,TRUE)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn binom_dist_range() {
    // microsoft documentation examples (0.084 and 0.524), mpmath 30 digits
    approx("BINOM.DIST.RANGE(60,0.75,48)", 0.0839749674290475);
    approx("BINOM.DIST.RANGE(60,0.75,45,50)", 0.5236297934718872);
    // mpmath 30 digits
    approx("BINOM.DIST.RANGE(100,0.3,20,40)", 0.978614384517419);
    // mpmath 30 digits
    approx("BINOM.DIST.RANGE(1000,0.5,400,600)", 0.9999999998198318);
    approx("BINOM.DIST.RANGE(60.9,0.75,48.5,48.9)", 0.0839749674290475);
    check(&[
        ("BINOM.DIST.RANGE(60,0.75,61)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(60,0.75,50,45)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(60,0.75,45,61)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(60,1.2,4)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(-1,0.5,0)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(60,0.75,-1)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(\"x\",0.75,4)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn binom_inv_returns_the_smallest_count_reaching_alpha() {
    // microsoft documentation example
    approx("BINOM.INV(6,0.5,0.75)", 4.0);
    approx("CRITBINOM(6,0.5,0.75)", 4.0);
    // mpmath 30 digits: smallest k with cdf(k) >= alpha
    approx("BINOM.INV(100,0.3,0.5)", 30.0);
    // mpmath 30 digits
    approx("BINOM.INV(1000,0.5,0.975)", 531.0);
    // mpmath 30 digits
    approx("BINOM.INV(50,0.1,0.01)", 1.0);
    // cdf(2) of six fair trials is exactly 22/64, which already reaches alpha
    approx("BINOM.INV(6,0.5,0.34375)", 2.0);
    approx("BINOM.INV(6,0.5,0.34376)", 3.0);
    approx("BINOM.INV(6.9,0.5,0.75)", 4.0);
    approx("BINOM.INV(6,0.5,0)", 0.0);
    check(&[
        ("BINOM.INV(6,0.5,1.5)", e(ErrorValue::Num)),
        ("BINOM.INV(6,0.5,-0.1)", e(ErrorValue::Num)),
        ("BINOM.INV(6,1.5,0.5)", e(ErrorValue::Num)),
        ("BINOM.INV(-1,0.5,0.5)", e(ErrorValue::Num)),
        ("CRITBINOM(6,0.5,1.5)", e(ErrorValue::Num)),
        ("BINOM.INV(\"x\",0.5,0.5)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn negbinom_dist() {
    // microsoft documentation examples (0.3135141 and 0.0550487), mpmath 30 digits
    approx("NEGBINOM.DIST(10,5,0.25,TRUE)", 0.3135140584781766);
    approx("NEGBINOM.DIST(10,5,0.25,FALSE)", 0.05504866037517786);
    approx("NEGBINOMDIST(10,5,0.25)", 0.05504866037517786);
    // mpmath 30 digits
    approx("NEGBINOM.DIST(20,3,0.1,TRUE)", 0.40804326879851816);
    // mpmath 30 digits
    approx("NEGBINOM.DIST(20,3,0.1,FALSE)", 0.028084207210421507);
    // mpmath 30 digits
    approx("NEGBINOM.DIST(100,50,0.4,TRUE)", 0.9611326191310378);
    // mpmath 30 digits
    approx("NEGBINOM.DIST(100,50,0.4,FALSE)", 0.005556716421281498);
    approx("NEGBINOM.DIST(10.7,5.9,0.25,FALSE)", 0.05504866037517786);
    check(&[
        ("NEGBINOM.DIST(10,0,0.25,TRUE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(-1,5,0.25,TRUE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(10,5,1.1,TRUE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(10,5,-0.1,FALSE)", e(ErrorValue::Num)),
        ("NEGBINOMDIST(10,0,0.25)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(\"x\",5,0.25,TRUE)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn hypgeom_dist() {
    // microsoft documentation examples (0.4654 and 0.3633), mpmath 30 digits
    approx("HYPGEOM.DIST(1,4,8,20,TRUE)", 0.46542827657378744);
    approx("HYPGEOM.DIST(1,4,8,20,FALSE)", 0.3632610939112487);
    approx("HYPGEOMDIST(1,4,8,20)", 0.3632610939112487);
    // mpmath 30 digits
    approx("HYPGEOM.DIST(45,100,500,1000,FALSE)", 0.04833670252843387);
    // mpmath 30 digits
    approx("HYPGEOM.DIST(45,100,500,1000,TRUE)", 0.17140977008783606);
    // mpmath 30 digits
    approx("HYPGEOM.DIST(30,60,70,200,TRUE)", 0.9988227785763978);
    approx("HYPGEOMDIST(45,100,500,1000)", 0.04833670252843387);
    approx("HYPGEOM.DIST(1.9,4.9,8.9,20.9,FALSE)", 0.3632610939112487);
    check(&[
        ("HYPGEOM.DIST(5,4,8,20,TRUE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(-1,4,8,20,TRUE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(0,15,10,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1,0,8,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1,21,8,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1,4,0,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1,4,21,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(0,4,8,0,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOMDIST(5,4,8,20)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(\"x\",4,8,20,TRUE)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn poisson_dist() {
    // microsoft documentation examples (0.124652 and 0.084224), mpmath 30 digits
    approx("POISSON.DIST(2,5,TRUE)", 0.12465201948308115);
    approx("POISSON.DIST(2,5,FALSE)", 0.08422433748856833);
    approx("POISSON(2,5,TRUE)", 0.12465201948308115);
    // mpmath 30 digits, regularized upper incomplete gamma
    approx("POISSON.DIST(150,140.5,TRUE)", 0.8017355898041434);
    // mpmath 30 digits
    approx("POISSON.DIST(150,140.5,FALSE)", 0.023778284438706637);
    // mpmath 30 digits
    approx("POISSON.DIST(3,0.01,TRUE)", 0.9999999995866529);
    approx("POISSON.DIST(2.9,5,FALSE)", 0.08422433748856833);
    approx("POISSON.DIST(3,0,TRUE)", 1.0);
    approx("POISSON.DIST(0,0,FALSE)", 1.0);
    check(&[
        ("POISSON.DIST(-1,5,TRUE)", e(ErrorValue::Num)),
        ("POISSON.DIST(2,-1,TRUE)", e(ErrorValue::Num)),
        ("POISSON(-1,5,FALSE)", e(ErrorValue::Num)),
        ("POISSON.DIST(\"a\",5,TRUE)", e(ErrorValue::Value)),
    ]);
}

#[test]
fn expon_and_weibull() {
    // microsoft documentation examples, mpmath 30 digits
    approx("EXPON.DIST(0.2,10,TRUE)", 0.8646647167633873);
    approx("EXPON.DIST(0.2,10,FALSE)", 1.353352832366127);
    approx("EXPONDIST(0.2,10,TRUE)", 0.8646647167633873);
    // mpmath 30 digits: the small-x cdf keeps its relative precision
    approx_in(
        &data_workbook(),
        "EXPON.DIST(1E-10,2,TRUE)",
        1.9999999998e-10,
        1e-22,
    );
    // mpmath 30 digits
    approx("EXPON.DIST(3,0.5,FALSE)", 0.11156508007421491);
    // microsoft documentation examples (0.929581 and 0.035589), mpmath 30 digits
    approx("WEIBULL.DIST(105,20,100,TRUE)", 0.9295813900692768);
    approx("WEIBULL.DIST(105,20,100,FALSE)", 0.03558886402450438);
    approx("WEIBULL(105,20,100,TRUE)", 0.9295813900692768);
    // mpmath 30 digits
    approx("WEIBULL.DIST(2.5,1.5,3,TRUE)", 0.5326728701282738);
    // mpmath 30 digits
    approx("WEIBULL.DIST(2.5,1.5,3,FALSE)", 0.21330467563740899);
    check(&[
        ("EXPON.DIST(-1,10,TRUE)", e(ErrorValue::Num)),
        ("EXPON.DIST(1,0,TRUE)", e(ErrorValue::Num)),
        ("EXPONDIST(1,-1,FALSE)", e(ErrorValue::Num)),
        ("EXPON.DIST(\"x\",10,TRUE)", e(ErrorValue::Value)),
        ("WEIBULL.DIST(-1,20,100,TRUE)", e(ErrorValue::Num)),
        ("WEIBULL.DIST(1,0,100,TRUE)", e(ErrorValue::Num)),
        ("WEIBULL.DIST(1,20,0,TRUE)", e(ErrorValue::Num)),
        ("WEIBULL(-1,20,100,FALSE)", e(ErrorValue::Num)),
        ("WEIBULL.DIST(1,20,100,\"x\")", e(ErrorValue::Value)),
    ]);
}

#[test]
fn fisher_standardize_and_permutations() {
    // microsoft documentation examples, mpmath 30 digits
    approx("FISHER(0.75)", 0.9729550745276566);
    approx("FISHERINV(0.972955)", 0.7499999673941484);
    // mpmath 30 digits
    approx("FISHER(-0.3)", -0.3095196042031117);
    // mpmath 30 digits
    approx("FISHERINV(2.5)", 0.9866142981514303);
    // microsoft documentation example
    approx("STANDARDIZE(42,40,1.5)", 4.0 / 3.0);
    approx("STANDARDIZE(-1,2,4)", -0.75);
    // microsoft documentation examples
    approx("PERMUT(100,3)", 970200.0);
    approx("PERMUT(3,2)", 6.0);
    approx("PERMUT(30,10)", 109027350432000.0);
    approx("PERMUT(5.9,2.9)", 20.0);
    approx("PERMUT(4,0)", 1.0);
    // microsoft documentation examples
    approx("PERMUTATIONA(3,2)", 9.0);
    approx("PERMUTATIONA(2,2)", 4.0);
    approx("PERMUTATIONA(10.9,3.9)", 1000.0);
    check(&[
        ("FISHER(1)", e(ErrorValue::Num)),
        ("FISHER(-1)", e(ErrorValue::Num)),
        ("FISHER(\"x\")", e(ErrorValue::Value)),
        ("FISHERINV(\"x\")", e(ErrorValue::Value)),
        ("STANDARDIZE(42,40,0)", e(ErrorValue::Num)),
        ("STANDARDIZE(42,40,-1)", e(ErrorValue::Num)),
        ("PERMUT(0,0)", e(ErrorValue::Num)),
        ("PERMUT(3,4)", e(ErrorValue::Num)),
        ("PERMUT(5,-1)", e(ErrorValue::Num)),
        ("PERMUT(200,200)", e(ErrorValue::Num)),
        ("PERMUT(\"x\",2)", e(ErrorValue::Value)),
        ("PERMUTATIONA(-1,2)", e(ErrorValue::Num)),
        ("PERMUTATIONA(2,-1)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn pearson_and_steyx() {
    // microsoft documentation examples, mpmath 30 digits
    approx("PEARSON({9,7,5,3,1},{10,6,1,5,3})", 0.6993786061802353);
    approx(
        "STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
        3.3057189502100415,
    );
    // mpmath 30 digits
    approx(
        "PEARSON({1.5,2.25,3.75,4.5,6.125,7},{2,3,5,7,11,13})",
        0.9908071003724461,
    );
    // mpmath 30 digits
    approx(
        "STEYX({1.5,2.25,3.75,4.5,6.125,7},{2,3,5,7,11,13})",
        0.3238992347717331,
    );
    approx(
        "PEARSON({9,7,5,3,1},{10,6,1,5,3})-CORREL({9,7,5,3,1},{10,6,1,5,3})",
        0.0,
    );
    check(&[
        ("STEYX({1,2},{3,4})", e(ErrorValue::Div0)),
        ("STEYX({1,2,3},{4,4,4})", e(ErrorValue::Div0)),
        ("STEYX({1,2,3},{1,2})", e(ErrorValue::NA)),
        ("PEARSON({1,2,3},{1,2})", e(ErrorValue::NA)),
    ]);
}

#[test]
fn skew_p() {
    // microsoft documentation example (0.303193), mpmath 30 digits
    approx("SKEW.P(3,4,5,2,3,4,5,6,4,7)", 0.3031933393541437);
    // mpmath 30 digits
    approx("SKEW.P({1,2,4,8,16})", 0.8890481348169543);
    approx("SKEW.P(1,2,4,8,16,K3)", 0.8890481348169543);
    check(&[
        ("SKEW.P(1,2)", e(ErrorValue::Div0)),
        ("SKEW.P(1,1,1)", e(ErrorValue::Div0)),
        ("SKEW.P(1,2,\"x\")", e(ErrorValue::Value)),
    ]);
}

#[test]
fn variance_a_family_counts_text_and_logicals() {
    // microsoft documentation examples, mpmath 30 digits
    approx("VARA(F1:F10)", 754.2666666666667);
    approx("VARPA(F1:F10)", 678.84);
    approx("STDEVPA(F1:F10)", 26.054558142482477);
    // K1:K5 reads as {1, 1, 0, 4}: TRUE is one, text zero, the blank skipped
    approx("VARA(K1:K5)", 3.0);
    approx("VARPA(K1:K5)", 2.25);
    approx("STDEVPA(K1:K5)", 1.5);
    approx("VARA(1,TRUE,4)", 3.0);
    approx("VARPA(1,\"1\",4)", 2.0);
    check(&[
        ("VARA(5)", e(ErrorValue::Div0)),
        ("VARPA(M1:M3)", e(ErrorValue::Div0)),
        ("STDEVPA(M1:M3)", e(ErrorValue::Div0)),
        ("VARA(1,\"x\")", e(ErrorValue::Value)),
        ("VARPA(1,#N/A)", e(ErrorValue::NA)),
    ]);
}

#[test]
fn exclusive_percentile_and_quartile() {
    // microsoft documentation examples
    approx("PERCENTILE.EXC(B1:B9,0.25)", 2.5);
    // rank 0.37 * 10 = 3.7 between 3 and 6
    approx("PERCENTILE.EXC(B1:B9,0.37)", 5.1);
    approx("PERCENTILE.EXC(B1:B9,0.9)", 9.0);
    approx("PERCENTILE.EXC(B1:B9,0.1)", 1.0);
    // microsoft documentation examples
    approx("QUARTILE.EXC(C1:C11,1)", 15.0);
    approx("QUARTILE.EXC(C1:C11,3)", 43.0);
    approx("QUARTILE.EXC(C1:C11,2)", 40.0);
    approx("QUARTILE.EXC(C1:C11,1.9)", 15.0);
    check(&[
        ("PERCENTILE.EXC(B1:B9,0)", e(ErrorValue::Num)),
        ("PERCENTILE.EXC(B1:B9,0.01)", e(ErrorValue::Num)),
        ("PERCENTILE.EXC(B1:B9,2)", e(ErrorValue::Num)),
        ("PERCENTILE.EXC(B1:B9,0.95)", e(ErrorValue::Num)),
        ("PERCENTILE.EXC(M1:M3,0.5)", e(ErrorValue::Num)),
        ("PERCENTILE.EXC(B1:B9,\"x\")", e(ErrorValue::Value)),
        ("QUARTILE.EXC(C1:C11,0)", e(ErrorValue::Num)),
        ("QUARTILE.EXC(C1:C11,4)", e(ErrorValue::Num)),
        ("QUARTILE.EXC({1,2},1)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn percent_rank_truncates_to_significance() {
    // microsoft documentation examples: 3/9, 5/9, 6/9 and a quarter of the way
    // from 5/9 to 6/9, each cut (not rounded) to three digits
    approx("PERCENTRANK.INC(A1:A10,2)", 0.333);
    approx("PERCENTRANK.INC(A1:A10,4)", 0.555);
    approx("PERCENTRANK.INC(A1:A10,8)", 0.666);
    approx("PERCENTRANK.INC(A1:A10,5)", 0.583);
    approx("PERCENTRANK(A1:A10,8)", 0.666);
    approx("PERCENTRANK.INC(A1:A10,8,6)", 0.666666);
    approx("PERCENTRANK.INC(A1:A10,8,1)", 0.6);
    approx("PERCENTRANK.INC(A1:A10,1)", 0.0);
    approx("PERCENTRANK.INC(A1:A10,13)", 1.0);
    approx("PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10,11},8)", 0.7);
    approx("PERCENTRANK.INC(A1:A10,8,2.9)", 0.66);
    // microsoft documentation examples: 7/10, (3 + 2.43/3)/10 and the same at
    // one digit
    approx("PERCENTRANK.EXC(B1:B9,7)", 0.7);
    approx("PERCENTRANK.EXC(B1:B9,5.43)", 0.381);
    approx("PERCENTRANK.EXC(B1:B9,5.43,1)", 0.3);
    approx("PERCENTRANK.EXC(B1:B9,1)", 0.1);
    check(&[
        ("PERCENTRANK.INC(A1:A10,0)", e(ErrorValue::NA)),
        ("PERCENTRANK.INC(A1:A10,14)", e(ErrorValue::NA)),
        ("PERCENTRANK.INC(A1:A10,2,0)", e(ErrorValue::Num)),
        ("PERCENTRANK.INC(M1:M3,2)", e(ErrorValue::Num)),
        ("PERCENTRANK.EXC(B1:B9,10)", e(ErrorValue::NA)),
        ("PERCENTRANK.EXC(B1:B9,5,0)", e(ErrorValue::Num)),
        ("PERCENTRANK(A1:A10,\"x\")", e(ErrorValue::Value)),
    ]);
}

#[test]
fn mode_mult_spills_every_mode_down_a_column() {
    // microsoft documentation example
    check_array("MODE.MULT(J1:J12)", 3, 1, &[n(1.0), n(2.0), n(3.0)], 0.0);
    check_array("MODE.MULT({4,4,2,2,9})", 2, 1, &[n(4.0), n(2.0)], 0.0);
    check_array("MODE.MULT(5,1,5)", 1, 1, &[n(5.0)], 0.0);
    check_array("MODE.MULT({1,2,3})", 1, 1, &[e(ErrorValue::NA)], 0.0);
    check_array(
        "_xlfn.MODE.MULT(J1:J12)",
        3,
        1,
        &[n(1.0), n(2.0), n(3.0)],
        0.0,
    );
    check(&[
        ("INDEX(MODE.MULT(J1:J12),2)", n(2.0)),
        ("MODE.MULT(1,2,3)", e(ErrorValue::NA)),
    ]);
}

#[test]
fn prob() {
    // microsoft documentation examples
    approx("PROB(D1:D4,E1:E4,2)", 0.1);
    approx("PROB(D1:D4,E1:E4,1,3)", 0.8);
    approx("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},0.5,2.5)", 0.4);
    approx("PROB(D1:D4,E1:E4,3,1)", 0.0);
    approx("PROB(D1:D4,E1:E4,7)", 0.0);
    check(&[
        ("PROB({0,1},{0.5,0.6},0)", e(ErrorValue::Num)),
        ("PROB({0,1},{0,1},0)", e(ErrorValue::Num)),
        ("PROB({0,1},{-0.5,1.5},0)", e(ErrorValue::Num)),
        ("PROB({0,1,2},{0.5,0.5},0)", e(ErrorValue::NA)),
        ("PROB(D1:D4,E1:E3,0)", e(ErrorValue::NA)),
        ("PROB(D1:D4,E1:E4,\"x\")", e(ErrorValue::Value)),
    ]);
}

#[test]
fn z_test() {
    // microsoft documentation examples (0.090574 and 0.863043), mpmath 30 digits
    approx("Z.TEST(G1:G10,4)", 0.09057419685136377);
    approx("Z.TEST(G1:G10,6)", 0.8630433891295299);
    approx("ZTEST(G1:G10,4)", 0.09057419685136377);
    approx(
        "2*MIN(Z.TEST(G1:G10,4),1-Z.TEST(G1:G10,4))",
        2.0 * 0.09057419685136377,
    );
    // mpmath 30 digits
    approx("Z.TEST(G1:G10,4,2)", 0.04099516050019149);
    // mpmath 30 digits
    approx("Z.TEST(G1:G10,5.2,1.5)", 0.5834855531402393);
    check(&[
        ("Z.TEST(M1:M3,4)", e(ErrorValue::NA)),
        ("Z.TEST({5},4)", e(ErrorValue::Div0)),
        ("Z.TEST({5,5,5},4)", e(ErrorValue::Div0)),
        ("Z.TEST(G1:G10,4,0)", e(ErrorValue::Num)),
        ("Z.TEST(G1:G10,\"x\")", e(ErrorValue::Value)),
    ]);
}

/// least squares by householder QR: relative 1e-11 for the regression arrays.
const FIT: f64 = 1e-11;

#[test]
fn logest_and_growth() {
    // microsoft documentation examples, mpmath 30 digits
    check_array(
        "LOGEST(I1:I6,H1:H6,TRUE,FALSE)",
        1,
        2,
        &[n(1.4632756281161758), n(495.30477015872754)],
        FIT,
    );
    check_array(
        "GROWTH(I1:I6,H1:H6)",
        6,
        1,
        &[
            n(32618.2037735398),
            n(47729.42261474788),
            n(69841.30085621761),
            n(102197.07337883263),
            n(149542.48674004624),
            n(218821.8762145961),
        ],
        FIT,
    );
    check_array(
        "GROWTH(I1:I6,H1:H6,H7:H8)",
        2,
        1,
        &[n(320196.7183634732), n(468536.0541840495)],
        FIT,
    );
    // mpmath 30 digits: the statistics of the logarithmic fit
    let na = e(ErrorValue::NA);
    check_array(
        "LOGEST(I1:I6,H1:H6,TRUE,TRUE)",
        5,
        2,
        &[
            n(1.4632756281161758),
            n(495.30477015872754),
            n(0.0026334028914250613),
            n(0.035834282435717624),
            n(0.9998086197758176),
            n(0.011016314665073015),
            n(20896.80109942032),
            n(4.0),
            n(2.5360188299385666),
            n(0.00048543675519961113),
        ],
        1e-9,
    );
    // mpmath 30 digits: const FALSE pins b to one
    check_array(
        "LOGEST(I1:I6,H1:H6,FALSE)",
        1,
        2,
        &[n(2.3003927637437855), n(1.0)],
        FIT,
    );
    check_array(
        "GROWTH(I1:I6,H1:H6,H7:H8,FALSE)",
        2,
        1,
        &[n(1414600.728288447), n(3254137.2789414325)],
        FIT,
    );
    let (rows, cols, values) = arrayed(&data_workbook(), "LOGEST(I1:I6,H1:H6,FALSE,TRUE)");
    assert_eq!((rows, cols), (5, 2));
    assert_eq!(values[3], na);
    check_array("LOGEST({1,2,0,4})", 1, 1, &[e(ErrorValue::Num)], 0.0);
    check_array("GROWTH({1,-2,3,4})", 1, 1, &[e(ErrorValue::Num)], 0.0);
    approx_in(
        &data_workbook(),
        "INDEX(LOGEST(I1:I6,H1:H6),1,2)",
        495.30477015872754,
        FIT,
    );
    approx_in(
        &data_workbook(),
        "_xlfn.GROWTH(I1:I6,H1:H6,17)",
        320196.7183634732,
        FIT,
    );
}

#[test]
fn scalar_functions_lift_over_arrays() {
    check_array(
        "BINOM.DIST({5,6},10,0.5,FALSE)",
        1,
        2,
        &[n(0.24609375), n(0.205078125)],
        1e-12,
    );
    // mpmath 30 digits
    check_array(
        "POISSON.DIST({2;150},{5;140.5},FALSE)",
        2,
        1,
        &[n(0.08422433748856833), n(0.023778284438706637)],
        1e-12,
    );
    check_array(
        "BINOM.INV(6,0.5,{0.75,0.34375,2})",
        1,
        3,
        &[n(4.0), n(2.0), e(ErrorValue::Num)],
        0.0,
    );
    check_array(
        "PERMUT({3;100},{2,3})",
        2,
        2,
        &[n(6.0), n(6.0), n(9900.0), n(970200.0)],
        0.0,
    );
    check_array(
        "FISHER({0.75,1})",
        1,
        2,
        &[n(0.9729550745276566), e(ErrorValue::Num)],
        1e-12,
    );
    check_array(
        "HYPGEOM.DIST({0,1},4,8,20,FALSE)",
        1,
        2,
        &[n(495.0 / 4845.0), n(0.3632610939112487)],
        1e-12,
    );
    check_array(
        "STANDARDIZE(D1:D4,1,2)",
        4,
        1,
        &[n(-0.5), n(0.0), n(0.5), n(1.0)],
        1e-12,
    );
}

#[test]
fn range_arguments_stay_whole_in_arrays() {
    check_array(
        "PERCENTRANK.INC(A1:A10,{2,4,8})",
        1,
        3,
        &[n(0.333), n(0.555), n(0.666)],
        1e-12,
    );
    check_array(
        "PERCENTRANK.INC(A1:A10,8,{1;2})",
        2,
        1,
        &[n(0.6), n(0.66)],
        1e-12,
    );
    check_array(
        "PERCENTILE.EXC(B1:B9,{0.25,0.01})",
        1,
        2,
        &[n(2.5), e(ErrorValue::Num)],
        1e-12,
    );
    check_array(
        "QUARTILE.EXC(C1:C11,{1;3})",
        2,
        1,
        &[n(15.0), n(43.0)],
        1e-12,
    );
    check_array(
        "PROB(D1:D4,E1:E4,{0,1},{1,3})",
        1,
        2,
        &[n(0.5), n(0.8)],
        1e-12,
    );
    // mpmath 30 digits
    check_array(
        "Z.TEST(G1:G10,{4,6})",
        1,
        2,
        &[n(0.09057419685136377), n(0.8630433891295299)],
        1e-12,
    );
    check_array("VARA(K1:K5)", 1, 1, &[n(3.0)], 1e-12);
    check_array(
        "STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
        1,
        1,
        &[n(3.3057189502100415)],
        1e-12,
    );
    check_array(
        "SKEW.P({3,4,5,2,3,4,5,6,4,7})",
        1,
        1,
        &[n(0.3031933393541437)],
        1e-12,
    );
    check_array(
        "PEARSON({9,7,5,3,1},{10,6,1,5,3})",
        1,
        1,
        &[n(0.6993786061802353)],
        1e-12,
    );
}

fn relative(src: &str, want: f64, tol: f64) {
    match eval_in(&data_workbook(), src) {
        CellValue::Number { value } => assert!(
            (value - want).abs() <= tol * want.abs(),
            "formula {src:?}: {value} != {want}"
        ),
        other => panic!("formula {src:?}: expected number, got {other:?}"),
    }
}

#[test]
fn huge_arguments_terminate() {
    // mpmath 30 digits: 1 / sqrt(2 pi 1e308), the stirling term below f64
    relative(
        "POISSON.DIST(1E307*10,1E307*10,FALSE)",
        3.989422804014327e-155,
        1e-12,
    );
    // mpmath 30 digits: 1/2 + pmf(n/2)/2 by symmetry
    relative(
        "BINOM.DIST(2000000000,4000000000,0.5,TRUE)",
        0.5000063078313046,
        1e-12,
    );
    // mpmath 30 digits: I_1/2(s, s+1) = 1/2 + 2^-2s / (s B(s, s))
    relative(
        "NEGBINOM.DIST(2000000000,2000000000,0.5,TRUE)",
        0.5000063078313046,
        1e-12,
    );
    // mpmath 30 digits
    relative(
        "BINOM.DIST.RANGE(4000000000,0.5,2000000000)",
        1.2615662609312322e-05,
        1e-12,
    );
    // mpmath 30 digits
    relative(
        "BINOM.DIST.RANGE(4000000000,0.5,2000300000)",
        3.6112562937369606e-25,
        1e-12,
    );
    relative("PERMUT(1E16,1)", 1e16, 1e-15);
    // mpmath 30 digits
    relative("PERMUT(2E16,4)", 1.5999999999999996e65, 1e-12);
    approx("BINOM.INV(1E16+2,1,0.5)", 10000000000000002.0);
    approx("BINOM.INV(1E16+2,0,0.5)", 0.0);
    check(&[
        ("BINOM.DIST(1E307*10,1.05E307*10,0.95,FALSE)", n(0.0)),
        ("POISSON.DIST(1E307*10,1E307*10,TRUE)", e(ErrorValue::Num)),
        ("BINOM.INV(1E307*10,0.9,0.9)", e(ErrorValue::Num)),
        ("BINOM.INV(4E16,0.5,0.5)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1E16,2E16,2E16,4E16,TRUE)", e(ErrorValue::Num)),
        ("BINOM.DIST(5E15,1E16,0.5,TRUE)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn negative_fractional_counts_are_num_errors() {
    check(&[
        ("BINOM.DIST(-0.5,10,0.5,FALSE)", e(ErrorValue::Num)),
        ("BINOM.DIST(1,-0.5,0.5,FALSE)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(10,0.5,-0.5)", e(ErrorValue::Num)),
        ("BINOM.DIST.RANGE(-0.5,0.5,0)", e(ErrorValue::Num)),
        ("BINOM.INV(-0.5,0.5,0.5)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(-0.5,5,0.25,FALSE)", e(ErrorValue::Num)),
        ("NEGBINOMDIST(-0.5,5,0.25)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(-0.5,4,8,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOMDIST(-0.5,4,8,20)", e(ErrorValue::Num)),
        ("POISSON.DIST(-0.5,5,FALSE)", e(ErrorValue::Num)),
        ("PERMUT(5,-0.5)", e(ErrorValue::Num)),
        ("PERMUTATIONA(-0.5,2)", e(ErrorValue::Num)),
        ("PERMUTATIONA(2,-0.5)", e(ErrorValue::Num)),
    ]);
    approx("BINOM.DIST(0.5,10,0.5,FALSE)", 0.0009765625);
    approx("POISSON.DIST(0.9,5,FALSE)", (-5.0f64).exp());
}

/// excel requires alpha below one.
#[test]
fn binom_inv_rejects_an_alpha_of_one() {
    check(&[
        // Excel 365 for Mac result
        ("BINOM.INV(1000,0.001,1)", e(ErrorValue::Num)),
        // Excel 365 for Mac result
        ("CRITBINOM(20,0.1,1)", e(ErrorValue::Num)),
        ("BINOM.INV(6,0.5,1)", e(ErrorValue::Num)),
    ]);
}

/// a skipped known_x fits against 1, 2, 3, ... as GROWTH does, and a skipped
/// const keeps the intercept.
#[test]
fn logest_with_skipped_known_x() {
    check_array("LOGEST({2,4,8},,FALSE)", 1, 2, &[n(2.0), n(1.0)], 1e-12);
    check_array("LOGEST({2,4,8},,TRUE)", 1, 2, &[n(2.0), n(1.0)], 1e-12);
    check_array("LOGEST({3,6,12},,,FALSE)", 1, 2, &[n(2.0), n(1.5)], 1e-12);
    check_array("LINEST({3,5,7},,,FALSE)", 1, 2, &[n(2.0), n(1.0)], 1e-12);
    check_array("LINEST({2,4,6},,FALSE)", 1, 2, &[n(2.0), n(0.0)], 1e-12);
    check_array(
        "LOGEST({1E200,1E201,1E202},,TRUE)",
        1,
        2,
        &[n(10.0), n(1e199)],
        1e-12,
    );
    check_array(
        "LOGEST({1E-200;1E-199;1E-198})",
        1,
        2,
        &[n(10.0), n(1e-201)],
        1e-12,
    );
    relative("INDEX(LOGEST({1E-200;1E-199;1E-198}),1,2)", 1e-201, 1e-12);
}

/// STEYX keeps excel's textbook formula, cancellation included; sums that
/// overflow are `#NUM!` rather than a zero.
#[test]
fn steyx_follows_the_textbook_formula() {
    check(&[
        // Excel 365 for Mac result
        ("STEYX({100000000,200000001,300000000},{1,2,3})", n(0.0)),
        (
            "STEYX({1E200,3E200,2E200,5E200},{1,2,3,4})",
            e(ErrorValue::Num),
        ),
    ]);
}

/// WEIBULL.DIST evaluates the density directly, so terms that overflow are
/// `#NUM!` as in excel.
#[test]
fn weibull_density_overflow_is_a_num_error() {
    check(&[
        // Excel 365 for Mac result
        ("WEIBULL.DIST(100,200,1,FALSE)", e(ErrorValue::Num)),
        ("WEIBULL(100,200,1,FALSE)", e(ErrorValue::Num)),
    ]);
}

/// functions that only exist in array form, nested as an argument of another
/// function, in a plain and an array formula. LOGEST's slope here is
/// exp(ln 2), one unit in the last place below 2, so it is rounded before it
/// is written as text.
#[test]
fn array_form_functions_nest_as_arguments() {
    let wb = data_workbook();
    let cases = [
        (
            "TEXTJOIN(\",\",TRUE,ROUND(LOGEST({2,4,8},,FALSE),12))",
            t("2,1"),
        ),
        ("COLUMNS(LOGEST({2,4,8},,FALSE))", n(2.0)),
        ("SUM(LOGEST({2,4,8},,FALSE))", n(3.0)),
        ("VARA(LOGEST({2,4,8},,FALSE))", n(0.5)),
        ("SUM(LINEST({3,5,7},{1,2,3}))", n(3.0)),
        ("SUM(TREND({1,2,3}))", n(6.0)),
        ("SUM(GROWTH({2,4,8}))", n(14.0)),
        ("SUM(MINVERSE({2,0;0,4}))", n(0.75)),
        ("STDEVPA(MUNIT(2))", n(0.5)),
        ("SUM(LET(a,LOGEST({2,4,8},,FALSE),a*2))", n(6.0)),
        ("SUM(MAP(LOGEST({2,4,8},,FALSE),LAMBDA(v,v*2)))", n(6.0)),
        ("TEXTJOIN(\",\",TRUE,MODE.MULT({1,1,2,2,3}))", t("1,2")),
    ];
    for (src, want) in cases {
        let got = eval_in(&wb, src);
        let (_, _, arrayed) = arrayed(&wb, src);
        for value in [got, arrayed[0].clone()] {
            match (&value, &want) {
                (CellValue::Number { value }, CellValue::Number { value: want }) => {
                    assert!((value - want).abs() <= 1e-12, "{src}: {value} != {want}")
                }
                _ => assert_eq!(value, want, "{src}"),
            }
        }
    }
}

/// the probability loops charge the recalculation budget per step, so many
/// long sums in one array run out of budget with `#NUM!` instead of stalling,
/// while one long search still fits: n is odd, so by symmetry the cdf at
/// (n - 1) / 2 is exactly one half.
#[test]
fn probability_loops_are_charged_to_the_budget() {
    let wb = data_workbook();
    approx_in(&wb, "BINOM.DIST.RANGE(100000,0.5,0,100000)", 1.0, 1e-9);
    approx_in(
        &wb,
        "BINOM.INV(39999999999,0.5,0.4999999)",
        19999999999.0,
        1e-15,
    );
    approx_in(
        &wb,
        "BINOM.INV(39999999999,0.5,0.5000001)",
        20000000000.0,
        1e-15,
    );
    for src in [
        "BINOM.DIST.RANGE(100000,0.5,0,100000-SEQUENCE(1,200))",
        "POISSON.DIST(1E9+SEQUENCE(1,100),1E9,TRUE)",
        "BINOM.DIST(2E9+SEQUENCE(1,100),4E9,0.5,TRUE)",
        "NEGBINOM.DIST(2E9+SEQUENCE(1,100),2E9,0.5,TRUE)",
        "HYPGEOM.DIST(1E9+SEQUENCE(1,200),2E9,4E9,8E9,TRUE)",
    ] {
        let (_, _, values) = arrayed(&wb, src);
        assert!(values.contains(&e(ErrorValue::Num)), "{src}: {values:?}");
    }
}
