//! database functions over microsoft's orchard example: criteria in A1:F3,
//! the database in A4:E10. statistics are checked against 30-digit decimal
//! oracles.

use xlsx_calc::{rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

const RESULT: &str = "ZZ1";

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

fn text(value: &str) -> Cell {
    Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Cell::default()
    }
}

fn formula(source: &str) -> Cell {
    Cell {
        formula: Some(source.into()),
        ..Cell::default()
    }
}

/// microsoft's orchard sheet: the criteria rows and the tree database.
fn orchard() -> Vec<(String, Cell)> {
    let mut cells: Vec<(String, Cell)> = [
        ("A1", text("Tree")),
        ("B1", text("Height")),
        ("C1", text("Age")),
        ("D1", text("Yield")),
        ("E1", text("Profit")),
        ("F1", text("Height")),
        ("A2", formula("\"=Apple\"")),
        ("B2", text(">10")),
        ("F2", text("<16")),
        ("A3", formula("\"=Pear\"")),
        ("A4", text("Tree")),
        ("B4", text("Height")),
        ("C4", text("Age")),
        ("D4", text("Yield")),
        ("E4", text("Profit")),
    ]
    .into_iter()
    .map(|(address, cell)| (address.to_string(), cell))
    .collect();
    let trees = [
        ("Apple", 18.0, 20.0, 14.0, 105.0),
        ("Pear", 12.0, 12.0, 10.0, 96.0),
        ("Cherry", 13.0, 14.0, 9.0, 105.0),
        ("Apple", 14.0, 15.0, 10.0, 75.0),
        ("Pear", 9.0, 8.0, 8.0, 76.8),
        ("Apple", 8.0, 9.0, 6.0, 45.0),
    ];
    for (i, (tree, height, age, yield_, profit)) in trees.into_iter().enumerate() {
        let row = 5 + i;
        cells.push((format!("A{row}"), text(tree)));
        cells.push((format!("B{row}"), number(height)));
        cells.push((format!("C{row}"), number(age)));
        cells.push((format!("D{row}"), number(yield_)));
        cells.push((format!("E{row}"), number(profit)));
    }
    cells
}

/// the orchard sheet with `extra` cells and `result` at [`RESULT`].
fn book_with(extra: &[(&str, Cell)], result: Cell) -> Workbook {
    let mut sheet = Sheet::new("Sheet1");
    for (address, cell) in orchard() {
        sheet.set_cell(CellRef::parse_a1(&address).unwrap(), cell);
    }
    for (address, cell) in extra {
        sheet.set_cell(CellRef::parse_a1(address).unwrap(), cell.clone());
    }
    sheet.set_cell(CellRef::parse_a1(RESULT).unwrap(), result);
    Workbook {
        sheets: vec![sheet],
        ..Workbook::default()
    }
}

fn result(book: &Workbook) -> CellValue {
    book.sheets[0]
        .cell(CellRef::parse_a1(RESULT).unwrap())
        .unwrap()
        .value
        .clone()
}

fn eval_with(extra: &[(&str, Cell)], source: &str) -> CellValue {
    let mut book = book_with(extra, formula(source));
    rebuild_and_recalc_all(&mut book, None);
    result(&book)
}

/// a formula carrying the value excel cached for it.
fn cached(source: &str, value: f64) -> Cell {
    Cell {
        formula: Some(source.into()),
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

/// set `address` to `cell` and recalculate what depends on it.
fn edit(book: &mut Workbook, graph: &mut xlsx_calc::graph::DepGraph, address: &str, cell: Cell) {
    let at = CellRef::parse_a1(address).unwrap();
    book.sheets[0].set_cell(at, cell);
    recalc_after(book, graph, &[(SheetId(0), at)], None);
}

fn value_with(extra: &[(&str, Cell)], source: &str) -> f64 {
    match eval_with(extra, source) {
        CellValue::Number { value } => value,
        other => panic!("{source}: expected a number, got {other:?}"),
    }
}

/// within relative `1e-12` of `want`.
fn close_with(extra: &[(&str, Cell)], source: &str, want: f64) {
    let got = value_with(extra, source);
    let scale = want.abs().max(f64::MIN_POSITIVE);
    assert!(
        (got - want).abs() <= 1e-12 * scale,
        "{source}: {got:e} vs {want:e}"
    );
}

fn close(source: &str, want: f64) {
    close_with(&[], source, want);
}

/// within `abs` of `want`: microsoft's examples print rounded results.
fn near(source: &str, want: f64, abs: f64) {
    let got = value_with(&[], source);
    assert!((got - want).abs() <= abs, "{source}: {got:e} vs {want:e}");
}

fn fails_with(extra: &[(&str, Cell)], source: &str, want: ErrorValue) {
    assert_eq!(
        eval_with(extra, source),
        CellValue::Error { value: want },
        "{source}"
    );
}

fn fails(source: &str, want: ErrorValue) {
    fails_with(&[], source, want);
}

#[test]
fn daverage_microsoft_examples() {
    close("DAVERAGE(A4:E10,\"Yield\",A1:B2)", 12.0);
    close("DAVERAGE(A4:E10,3,A4:E10)", 13.0);
}

#[test]
fn dcount_and_dcounta_microsoft_examples() {
    close("DCOUNT(A4:E10,\"Age\",A1:F2)", 1.0);
    close("DCOUNTA(A4:E10,\"Profit\",A1:F2)", 1.0);
}

#[test]
fn dget_microsoft_example_more_than_one_match_is_num() {
    fails("DGET(A4:E10,\"Yield\",A1:A3)", ErrorValue::Num);
}

#[test]
fn dmax_and_dmin_microsoft_examples() {
    close("DMAX(A4:E10,\"Profit\",A1:F3)", 96.0);
    close("DMIN(A4:E10,\"Profit\",A1:B2)", 75.0);
}

#[test]
fn dproduct_microsoft_examples() {
    close("DPRODUCT(A4:E10,\"Yield\",A1:F3)", 800.0);
    close("DPRODUCT(A4:E10,\"Yield\",A1:F2)", 10.0);
}

#[test]
fn dstdev_and_dstdevp_microsoft_examples() {
    near("DSTDEV(A4:E10,\"Yield\",A1:A3)", 2.97, 0.005);
    near("DSTDEVP(A4:E10,\"Yield\",A1:A3)", 2.65, 0.005);
    // decimal 30 digits: sqrt(8.8), sqrt(7.04), rounded to f64
    close("DSTDEV(A4:E10,\"Yield\",A1:A3)", 2.966_479_394_838_265_3);
    close("DSTDEVP(A4:E10,\"Yield\",A1:A3)", 2.653_299_832_284_319_8);
}

#[test]
fn dsum_microsoft_examples() {
    close("DSUM(A4:E10,\"Profit\",A1:A2)", 225.0);
    close("DSUM(A4:E10,\"Profit\",A1:F2)", 75.0);
}

#[test]
fn dvar_and_dvarp_microsoft_examples() {
    close("DVAR(A4:E10,\"Yield\",A1:A3)", 8.8);
    close("DVARP(A4:E10,\"Yield\",A1:A3)", 7.04);
}

/// a criteria row with every cell blank sets no condition, so every record
/// matches.
#[test]
fn an_empty_criteria_row_matches_every_record() {
    let blank = [("H1", text("Tree"))];
    close_with(&blank, "DCOUNT(A4:E10,,H1:H2)", 6.0);
    close_with(&blank, "DSUM(A4:E10,\"Profit\",H1:H2)", 502.8);
    // decimal 30 digits over all six records
    close_with(
        &blank,
        "DSTDEV(A4:E10,\"Yield\",H1:H2)",
        2.664_582_518_894_845_5,
    );
    close_with(&blank, "DVAR(A4:E10,\"Yield\",H1:H2)", 7.1);
    close_with(
        &blank,
        "DSTDEV(A4:E10,\"Profit\",H1:H2)",
        23.149_946_004_256_684,
    );
    close_with(&blank, "DVARP(A4:E10,\"Profit\",H1:H2)", 446.6);
    close_with(&blank, "DAVERAGE(A4:E10,\"Profit\",H1:H2)", 83.8);
}

#[test]
fn criteria_rows_are_alternatives() {
    close("DCOUNT(A4:E10,,A1:A3)", 5.0);
    close("DCOUNTA(A4:E10,,A1:A3)", 5.0);
    close("DSUM(A4:E10,\"Yield\",A1:A3)", 48.0);
}

/// the same label twice in one row puts two conditions on one column.
#[test]
fn one_column_twice_in_a_row_must_satisfy_both() {
    let band = [
        ("H1", text("Height")),
        ("I1", text("Height")),
        ("H2", text(">=12")),
        ("I2", text("<14")),
    ];
    close_with(&band, "DSUM(A4:E10,\"Profit\",H1:I2)", 201.0);
    close_with(&band, "DCOUNT(A4:E10,\"Height\",H1:I2)", 2.0);
}

#[test]
fn wildcards_and_the_leading_equals_sign() {
    let pattern = |criterion: &str| vec![("H1", text("Tree")), ("H2", text(criterion))];
    close_with(&pattern("A*e"), "DSUM(A4:E10,\"Profit\",H1:H2)", 225.0);
    close_with(&pattern("?ear"), "DSUM(A4:E10,\"Profit\",H1:H2)", 172.8);
    close_with(&pattern("=P*"), "DSUM(A4:E10,\"Profit\",H1:H2)", 172.8);
    close_with(&pattern("<>Apple"), "DCOUNTA(A4:E10,,H1:H2)", 3.0);
    close_with(&pattern("=?????"), "DCOUNTA(A4:E10,,H1:H2)", 3.0);
}

/// text without an operator asks for text that begins with it, the way a
/// criteria range reads; `=text` asks for the whole text.
#[test]
fn plain_text_criteria_match_by_prefix() {
    let pattern = |criterion: &str| vec![("H1", text("Tree")), ("H2", text(criterion))];
    close_with(&pattern("Ch"), "DSUM(A4:E10,\"Profit\",H1:H2)", 105.0);
    close_with(&pattern("ap"), "DCOUNTA(A4:E10,,H1:H2)", 3.0);
    close_with(&pattern("=Ch"), "DCOUNTA(A4:E10,,H1:H2)", 0.0);
    close_with(&pattern("=cherry"), "DCOUNTA(A4:E10,,H1:H2)", 1.0);
}

#[test]
fn numeric_criteria_compare_numbers() {
    let crit = [("H1", text("Age")), ("H2", number(15.0))];
    close_with(&crit, "DGET(A4:E10,\"Profit\",H1:H2)", 75.0);
    let crit = [("H1", text("Yield")), ("H2", text("<=9"))];
    close_with(&crit, "DSUM(A4:E10,\"Profit\",H1:H2)", 226.8);
}

#[test]
fn a_numeric_field_is_a_truncated_position() {
    close("DSUM(A4:E10,5,A1:A2)", 225.0);
    close("DSUM(A4:E10,5.9,A1:A2)", 225.0);
    fails("DSUM(A4:E10,0,A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,6,A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,-1,A1:A2)", ErrorValue::Value);
}

#[test]
fn labels_match_without_regard_to_case() {
    close("DSUM(A4:E10,\"PROFIT\",A1:A2)", 225.0);
    close("DSUM(A4:E10,\"yield\",A1:A2)", 30.0);
    let crit = [("H1", text("tREE")), ("H2", text("=pear"))];
    close_with(&crit, "DMAX(A4:E10,\"profit\",H1:H2)", 96.0);
}

#[test]
fn dget_returns_the_one_matching_field() {
    let crit = [("H1", text("Tree")), ("H2", text("=Cherry"))];
    close_with(&crit, "DGET(A4:E10,\"Yield\",H1:H2)", 9.0);
    assert_eq!(
        eval_with(&crit, "DGET(A4:E10,1,H1:H2)"),
        CellValue::Text {
            value: "Cherry".into()
        }
    );
    let crit = [("H1", text("Tree")), ("H2", text("=Plum"))];
    fails_with(&crit, "DGET(A4:E10,\"Yield\",H1:H2)", ErrorValue::Value);
}

#[test]
fn no_matching_record() {
    let crit = [("H1", text("Tree")), ("H2", text("=Plum"))];
    close_with(&crit, "DSUM(A4:E10,\"Profit\",H1:H2)", 0.0);
    close_with(&crit, "DCOUNT(A4:E10,,H1:H2)", 0.0);
    close_with(&crit, "DMAX(A4:E10,\"Profit\",H1:H2)", 0.0);
    close_with(&crit, "DMIN(A4:E10,\"Profit\",H1:H2)", 0.0);
    close_with(&crit, "DPRODUCT(A4:E10,\"Profit\",H1:H2)", 0.0);
    fails_with(&crit, "DAVERAGE(A4:E10,\"Profit\",H1:H2)", ErrorValue::Div0);
    fails_with(&crit, "DVARP(A4:E10,\"Profit\",H1:H2)", ErrorValue::Div0);
    fails_with(&crit, "DSTDEVP(A4:E10,\"Profit\",H1:H2)", ErrorValue::Div0);
}

#[test]
fn a_sample_statistic_of_one_record_is_div0() {
    let crit = [("H1", text("Tree")), ("H2", text("=Cherry"))];
    fails_with(&crit, "DSTDEV(A4:E10,\"Yield\",H1:H2)", ErrorValue::Div0);
    fails_with(&crit, "DVAR(A4:E10,\"Yield\",H1:H2)", ErrorValue::Div0);
    close_with(&crit, "DVARP(A4:E10,\"Yield\",H1:H2)", 0.0);
}

#[test]
fn counts_skip_or_take_text_by_function() {
    close("DCOUNT(A4:E10,\"Tree\",A1:A3)", 0.0);
    close("DCOUNTA(A4:E10,\"Tree\",A1:A3)", 5.0);
    close("DSUM(A4:E10,\"Tree\",A1:A3)", 0.0);
}

#[test]
fn invalid_arguments() {
    fails("DSUM(A4:E10,\"Weight\",A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,,A1:A2)", ErrorValue::Value);
    fails("DGET(A4:E10,,A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,TRUE,A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,NA(),A1:A2)", ErrorValue::NA);
    fails("DSUM({1,2;3,4},1,A1:A2)", ErrorValue::Value);
    fails("DSUM(A4:E10,\"Profit\",\"Tree\")", ErrorValue::Value);
    fails("DSUM(A4:E10,\"Profit\",A1:E1)", ErrorValue::Value);
}

/// a formula under a label that names no field is evaluated for each record,
/// its relative references moved from the first record down to that one.
#[test]
fn computed_criteria_follow_each_record() {
    let older = [
        ("H1", text("Older")),
        ("H2", formula("C5>AVERAGE($C$5:$C$10)")),
    ];
    close_with(&older, "DSUM(A4:E10,\"Profit\",H1:H2)", 285.0);
    let blank_label = [("H2", formula("C5>AVERAGE($C$5:$C$10)"))];
    close_with(&blank_label, "DCOUNT(A4:E10,,H1:H2)", 3.0);
    let with_tree = [
        ("H1", text("Tree")),
        ("I1", text("Older")),
        ("H2", text("=Apple")),
        ("I2", formula("C5>AVERAGE($C$5:$C$10)")),
    ];
    close_with(&with_tree, "DSUM(A4:E10,\"Profit\",H1:I2)", 180.0);
    let ratio = [("H1", text("Ratio")), ("H2", formula("E5/D5>8"))];
    close_with(&ratio, "DCOUNT(A4:E10,,H1:H2)", 3.0);
}

/// under a field's own label the formula's value is an ordinary criterion:
/// TRUE asks for ages equal to TRUE, which no record has.
#[test]
fn a_formula_under_a_field_label_is_its_value() {
    let crit = [
        ("H1", text("Age")),
        ("H2", formula("C5>AVERAGE($C$5:$C$10)")),
    ];
    close_with(&crit, "DCOUNT(A4:E10,,H1:H2)", 0.0);
    let crit = [("H1", text("Age")), ("H2", formula("\">\"&13"))];
    close_with(&crit, "DCOUNT(A4:E10,,H1:H2)", 3.0);
}

/// the field lifts in an array formula, one answer per field.
#[test]
fn the_field_lifts_over_an_array() {
    close("SUM(DSUM(A4:E10,{\"Yield\",\"Profit\"},A1:A2))", 255.0);
}

#[test]
fn lowercase_names_resolve() {
    close("dsum(A4:E10,\"Profit\",A1:A2)", 225.0);
}

/// a computed criterion whose moved references leave the database reads
/// cells no dependency tracks, so the engine records a gap: the cell keeps
/// the value excel cached, before and after an edit there, and shows `#NAME?`
/// without one.
#[test]
fn a_computed_criterion_reading_past_the_database_keeps_the_cached_value() {
    let extra = [
        ("F5", number(0.0)),
        ("F6", number(1.0)),
        ("H1", text("Check")),
        ("H2", formula("F5>0")),
    ];
    let source = "DSUM(A4:E10,\"Profit\",H1:H2)";
    let mut book = book_with(&extra, cached(source, 96.0));
    let (mut graph, _) = rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 96.0 });
    edit(&mut book, &mut graph, "F6", number(0.0));
    assert_eq!(result(&book), CellValue::Number { value: 96.0 });
    fails_with(&extra, source, ErrorValue::Name);
}

/// the moved reference would read `F6` before its own formula ran.
#[test]
fn a_computed_criterion_reading_an_unordered_formula_keeps_the_cached_value() {
    let extra = [
        ("Z20", formula("12")),
        ("F5", number(18.0)),
        ("F6", formula("Z20")),
        ("H1", text("Check")),
        ("H2", formula("F5=12")),
    ];
    let source = "DSUM(A4:E10,\"Profit\",H1:H2)";
    let mut book = book_with(&extra, cached(source, 96.0));
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 96.0 });
}

#[test]
fn a_computed_criterion_reading_a_column_past_the_database_is_a_gap() {
    let extra = [("H1", text("Ratio")), ("H2", formula("E5/D5>8"))];
    fails_with(&extra, "DSUM(A4:D10,\"Yield\",H1:H2)", ErrorValue::Name);
    let mut book = book_with(&extra, cached("DSUM(A4:D10,\"Yield\",H1:H2)", 27.0));
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 27.0 });
}

/// moved references inside the database are covered by its dependency, so
/// an edit there recalculates the call.
#[test]
fn a_computed_criterion_inside_the_database_follows_edits() {
    let extra = [("H1", text("Ratio")), ("H2", formula("E5/D5>8"))];
    let mut book = book_with(&extra, formula("DCOUNT(A4:E10,,H1:H2)"));
    let (mut graph, _) = rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 3.0 });
    edit(&mut book, &mut graph, "E6", number(50.0));
    assert_eq!(result(&book), CellValue::Number { value: 2.0 });
}

/// every record a criteria row is tried on is charged, so a computed
/// criterion over a million records runs out of budget with `#NUM!`.
#[test]
fn large_databases_are_charged_to_the_budget() {
    let rows: Vec<(String, Cell)> = (2..=1001)
        .map(|row| (format!("H{row}"), formula("1=0")))
        .collect();
    let extra: Vec<(&str, Cell)> = rows
        .iter()
        .map(|(address, cell)| (address.as_str(), cell.clone()))
        .collect();
    fails_with(&extra, "DCOUNT(A2:A1048576,,H1:H1001)", ErrorValue::Num);
    let one = [("H2", formula("1=0"))];
    fails_with(
        &one,
        "SUM(DCOUNT(A2:A1048576,SEQUENCE(1000),H1:H2))",
        ErrorValue::Num,
    );
}

/// records past the sheet's used rows are blank: they are counted at once,
/// not read one by one.
#[test]
fn blank_records_past_the_used_rows_count_without_reading() {
    let all = [("H1", text("Tree"))];
    close_with(&all, "DCOUNTA(A4:E1048576,,H1:H2)", 1_048_572.0);
    close_with(&all, "DCOUNT(A4:E1048576,\"Yield\",H1:H2)", 6.0);
    fails_with(&all, "DGET(A4:E1048576,\"Yield\",H1:H2)", ErrorValue::Num);
    close("DSUM(A4:E1048576,\"Profit\",A1:A2)", 225.0);
    close("DCOUNT(A4:E1048576,,A1:A3)", 5.0);
    let cherry = [("H1", text("Tree")), ("H2", text("=Cherry"))];
    close_with(&cherry, "DGET(A4:E1048576,\"Yield\",H1:H2)", 9.0);
    let blank = [("H1", text("Tree")), ("H2", text("="))];
    close_with(&blank, "DCOUNTA(A4:E1048576,,H1:H2)", 1_048_566.0);
}

/// a full-height database keeps every record: the blank ones past the used
/// rows match a blank or `=` criterion and are counted.
#[test]
fn full_height_databases_count_their_blank_records() {
    let all = [("G1", text("Tree")), ("H1", text("Tree"))];
    close_with(&all, "DCOUNTA(G1:G1048576,,H1:H2)", 1_048_575.0);
    close_with(&all, "DCOUNTA(G:G,,H1:H2)", 1_048_575.0);
    close_with(&all, "DCOUNT(G1:G1048576,,H1:H2)", 1_048_575.0);
    let equals = [
        ("G1", text("Tree")),
        ("H1", text("Tree")),
        ("H2", text("=")),
    ];
    close_with(&equals, "DCOUNTA(G1:G1048576,,H1:H2)", 1_048_575.0);
}

#[test]
fn dget_over_a_full_height_database_sees_the_blank_matches() {
    let mut extra = vec![
        ("G1", text("Tree")),
        ("H1", text("Tree")),
        ("H2", text("=")),
    ];
    let apples: Vec<String> = (2..=10).map(|row| format!("G{row}")).collect();
    extra.extend(
        apples
            .iter()
            .map(|address| (address.as_str(), text("Apple"))),
    );
    fails_with(&extra, "DGET(G1:G1048576,\"Tree\",H1:H2)", ErrorValue::Num);
}

/// a computed criterion over a full-height database is tried on every
/// record, so it runs out of budget.
#[test]
fn computed_criteria_over_full_height_databases_are_charged() {
    let never = [("H2", formula("1=0"))];
    fails_with(
        &never,
        "SUM(DCOUNT(A1:A1048576,SEQUENCE(1000),H1:H2))",
        ErrorValue::Num,
    );
    let blank = [("G1", text("Tree")), ("H2", formula("ISBLANK(G2)"))];
    fails_with(&blank, "DCOUNT(G1:G1048576,,H1:H2)", ErrorValue::Num);
}

/// the blank rows a full-height criteria range holds past its used cells
/// each match every record.
#[test]
fn a_full_height_criteria_range_has_blank_rows() {
    let crit = [("H1", text("Tree")), ("H2", text("=Apple"))];
    close_with(&crit, "DCOUNTA(A4:E10,,H:H)", 6.0);
}

/// a range join's rectangle can reach past the database whatever its ends.
#[test]
fn a_computed_criterion_with_a_range_join_is_a_gap() {
    let extra = [
        ("F5", number(0.0)),
        ("F6", number(1.0)),
        ("H2", formula("SUM(($F$5):C5)>0")),
    ];
    let source = "DSUM(A4:E10,\"Profit\",H1:H2)";
    fails_with(&extra, source, ErrorValue::Name);
    let mut book = book_with(&extra, cached(source, 50.0));
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 50.0 });
    let index = [("H2", formula("SUM($F$5:INDEX(C5:C5,1))>0"))];
    fails_with(&index, source, ErrorValue::Name);
}

/// a reference a function builds is not bounded by its arguments.
#[test]
fn a_computed_criterion_building_a_reference_is_a_gap() {
    let source = "DSUM(A4:E10,\"Profit\",H1:H2)";
    for criterion in [
        "OFFSET(E5,0,1)>0",
        "INDIRECT(\"F\"&ROW(E5))>0",
        "SUM(CHOOSE(1,F5,E5))>0",
    ] {
        let extra = [
            ("F5", number(0.0)),
            ("F6", number(1.0)),
            ("H2", formula(criterion)),
        ];
        fails_with(&extra, source, ErrorValue::Name);
    }
}

/// `C6:$C$5` is stored top to bottom with its anchors on the wrong ends, so
/// a relative top over an anchored bottom cannot be moved faithfully.
#[test]
fn a_range_with_a_relative_top_and_an_anchored_bottom_is_a_gap() {
    let upside_down = [("H2", formula("SUM(C6:$C$5)>25"))];
    fails_with(&upside_down, "DCOUNT(A4:E9,,H1:H2)", ErrorValue::Name);
    let reaching = [("H2", formula("SUM(C10:$C$5)>0"))];
    fails_with(&reaching, "DSUM(A4:E10,\"Profit\",H1:H2)", ErrorValue::Name);
}

/// a computed criterion is evaluated at its own cell, moved down with the
/// record, so COLUMN() and ROW() describe the criteria cell.
#[test]
fn computed_criteria_evaluate_at_the_criteria_cell() {
    let column = [("H2", formula("COLUMN()=8"))];
    close_with(&column, "DSUM(A4:E10,\"Profit\",H1:H2)", 502.8);
    let row = [("H2", formula("ROW()<=3"))];
    close_with(&row, "DSUM(A4:E10,\"Profit\",H1:H2)", 201.0);
}

/// a function the engine lacks leaves the criterion unknown: the call is
/// `#NAME?`, or keeps the value excel cached, never a partial total.
#[test]
fn an_unsupported_function_in_a_computed_criterion_is_a_gap() {
    let extra = [("H2", formula("NOSUCHFUNCTION(A5)>0"))];
    let source = "DSUM(A4:E10,\"Profit\",H1:H2)";
    fails_with(&extra, source, ErrorValue::Name);
    let mut book = book_with(&extra, cached(source, 42.0));
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(result(&book), CellValue::Number { value: 42.0 });
    let handled = [("H2", formula("IFERROR(NOSUCHFUNCTION(A5),FALSE)"))];
    close_with(&handled, source, 0.0);
}

fn huge(values: &[f64]) -> Vec<(&'static str, Cell)> {
    let mut cells = vec![("J1", text("N")), ("K1", text("N"))];
    for (address, value) in ["J2", "J3", "J4"].into_iter().zip(values) {
        cells.push((address, number(*value)));
    }
    cells
}

#[test]
fn huge_fields_overflow_to_num_not_a_wrong_number() {
    fails_with(
        &huge(&[1e200, 1e200]),
        "DPRODUCT(J1:J4,\"N\",K1:K2)",
        ErrorValue::Num,
    );
    fails_with(
        &huge(&[1e308, 1e308]),
        "DSUM(J1:J4,\"N\",K1:K2)",
        ErrorValue::Num,
    );
    close_with(
        &huge(&[1e200, 1e200, 1e200]),
        "DVAR(J1:J4,\"N\",K1:K2)",
        0.0,
    );
    close_with(
        &huge(&[1e200, 1e200, 1e200]),
        "DVARP(J1:J4,\"N\",K1:K2)",
        0.0,
    );
    close_with(
        &huge(&[1e200, -1e200, 1e-200]),
        "DMAX(J1:J4,\"N\",K1:K2)",
        1e200,
    );
    close_with(
        &huge(&[3e-308, -0.0, 1e-307]),
        "DSUM(J1:J4,\"N\",K1:K2)",
        3e-308 + 1e-307,
    );
    for (database, plain) in [
        ("DSTDEV(J1:J4,\"N\",K1:K2)", "STDEV(J2:J4)"),
        ("DVAR(J1:J4,\"N\",K1:K2)", "VAR(J2:J4)"),
        ("DAVERAGE(J1:J4,\"N\",K1:K2)", "AVERAGE(J2:J4)"),
    ] {
        for values in [[1e200, -1e200], [1e308, 1e308]] {
            assert_eq!(
                eval_with(&huge(&values), database),
                eval_with(&huge(&values), plain),
                "{database} over {values:?}"
            );
        }
    }
}

#[test]
fn huge_field_positions_are_value_errors() {
    for field in ["1E+200", "-0.5", "1E15+0.75", "2^53+1"] {
        fails(&format!("DSUM(A4:E10,{field},A1:A2)"), ErrorValue::Value);
    }
}

/// a database spanning every column of a million rows is cut to the used
/// columns and its blank records counted at once.
#[test]
fn full_sheet_databases_stay_bounded() {
    close("DCOUNTA(4:1048576,,A1:A2)", 3.0);
    close("DSUM(A4:XFD1048576,\"Profit\",A1:A3)", 397.8);
    let all = [("H1", text("Tree"))];
    close_with(&all, "DCOUNTA(A4:XFD1048576,,H1:H2)", 1_048_572.0);
}
