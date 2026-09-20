//! array evaluation: the dynamic-array builtins, elementwise lifting of scalar
//! functions and operators, and how a `t="array"` formula lays its result out.

use xlsx_calc::{
    EvalContext, Value, evaluate_array, parse_formula, rebuild_and_recalc_all, recalc_after,
};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ErrorValue, MAX_SPILL_CELLS, Sheet, SheetId, Workbook,
};

fn n(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn t(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn a1(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn range(reference: &str) -> CellRange {
    CellRange::parse_a1(reference).unwrap()
}

fn put(sheet: &mut Sheet, address: &str, value: CellValue) {
    sheet.set_cell(
        a1(address),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

/// A1:B4 is a small table: names in A, numbers in B, one repeat and one blank.
fn fixture() -> Workbook {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Data");
    for (index, name) in ["  pear ", "apple", "pear", "apple"].iter().enumerate() {
        put(&mut sheet, &format!("A{}", index + 1), t(name));
    }
    for (index, value) in [3.0, 1.0, 4.0, 2.0].iter().enumerate() {
        put(&mut sheet, &format!("B{}", index + 1), n(*value));
    }
    workbook.sheets.push(sheet);
    workbook
}

/// evaluate in array mode and describe the result as (rows, cols, values).
fn arrayed(formula: &str, workbook: &Workbook) -> (usize, usize, Vec<CellValue>) {
    let expression = parse_formula(formula).unwrap();
    let context = EvalContext::new(workbook, SheetId(0));
    match evaluate_array(&expression, &context) {
        Value::Scalar(value) => (1, 1, vec![value]),
        Value::Array(array) => {
            let mut values = Vec::new();
            for row in 0..array.rows() {
                for col in 0..array.cols() {
                    values.push(array.at(row, col));
                }
            }
            (array.rows(), array.cols(), values)
        }
    }
}

fn values(formula: &str, workbook: &Workbook) -> Vec<CellValue> {
    arrayed(formula, workbook).2
}

#[test]
fn filter_keeps_the_rows_its_condition_selects() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn._xlws.FILTER(A1:B4,B1:B4>2)", &workbook),
        (2, 2, vec![t("  pear "), n(3.0), t("pear"), n(4.0)])
    );
    assert_eq!(
        values("_xlfn._xlws.FILTER(A1:A4,B1:B4>9,\"none\")", &workbook),
        vec![t("none")]
    );
}

#[test]
fn sort_orders_rows_by_the_requested_column() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn._xlws.SORT(B1:B4)", &workbook),
        vec![n(1.0), n(2.0), n(3.0), n(4.0)]
    );
    assert_eq!(
        values("_xlfn._xlws.SORT(A1:B4,2,-1)", &workbook),
        vec![
            t("pear"),
            n(4.0),
            t("  pear "),
            n(3.0),
            t("apple"),
            n(2.0),
            t("apple"),
            n(1.0)
        ]
    );
}

#[test]
fn sortby_orders_by_a_separate_key() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn.SORTBY(A1:A4,B1:B4,-1)", &workbook),
        vec![t("pear"), t("  pear "), t("apple"), t("apple")]
    );
}

#[test]
fn unique_keeps_first_occurrences_and_can_demand_singletons() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn.UNIQUE(A1:A4)", &workbook),
        vec![t("  pear "), t("apple"), t("pear")]
    );
    assert_eq!(
        values("_xlfn.UNIQUE(A1:A4,FALSE,TRUE)", &workbook),
        vec![t("  pear "), t("pear")]
    );
}

/// the grouping the hashed dedup key has to reproduce.
#[test]
fn unique_groups_blanks_case_and_types_the_way_excel_does() {
    let mut workbook = fixture();
    let sheet = workbook.sheet_mut(SheetId(0)).unwrap();
    for (address, value) in [
        ("D1", n(0.0)),
        ("D2", CellValue::Empty),
        ("D3", t("Pear")),
        ("D4", t("pear")),
        ("D5", CellValue::Bool { value: false }),
        (
            "D6",
            CellValue::Error {
                value: ErrorValue::NA,
            },
        ),
        (
            "D7",
            CellValue::Error {
                value: ErrorValue::Div0,
            },
        ),
    ] {
        put(sheet, address, value);
    }
    assert_eq!(
        values("_xlfn.UNIQUE(D1:D7)", &workbook),
        vec![
            n(0.0),
            t("Pear"),
            CellValue::Bool { value: false },
            CellValue::Error {
                value: ErrorValue::NA
            },
            CellValue::Error {
                value: ErrorValue::Div0
            },
        ]
    );
}

#[test]
fn sequence_builds_a_rectangle_from_start_and_step() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn.SEQUENCE(2,3,0,5)", &workbook),
        (
            2,
            3,
            vec![n(0.0), n(5.0), n(10.0), n(15.0), n(20.0), n(25.0)]
        )
    );
}

#[test]
fn stacking_and_slicing_reshape_blocks() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn.HSTACK(B1:B2,B3:B4)", &workbook),
        (2, 2, vec![n(3.0), n(4.0), n(1.0), n(2.0)])
    );
    assert_eq!(
        arrayed("_xlfn.VSTACK(B1:B2,B3:B4)", &workbook),
        (4, 1, vec![n(3.0), n(1.0), n(4.0), n(2.0)])
    );
    assert_eq!(
        values("_xlfn.TAKE(B1:B4,2)", &workbook),
        vec![n(3.0), n(1.0)]
    );
    assert_eq!(values("_xlfn.DROP(B1:B4,-3)", &workbook), vec![n(3.0)]);
    assert_eq!(
        values("_xlfn.CHOOSECOLS(A1:B2,{2,1})", &workbook),
        vec![n(3.0), t("  pear "), n(1.0), t("apple")]
    );
    assert_eq!(
        arrayed("_xlfn.TOROW(A1:B1)", &workbook),
        (1, 2, vec![t("  pear "), n(3.0)])
    );
    assert_eq!(
        arrayed("TRANSPOSE(A1:B1)", &workbook),
        (2, 1, vec![t("  pear "), n(3.0)])
    );
}

#[test]
fn expand_pads_to_the_requested_size() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn.EXPAND(B1:B2,2,2,\"-\")", &workbook),
        (2, 2, vec![n(3.0), t("-"), n(1.0), t("-")])
    );
}

/// a single-argument scalar function applied to a block runs once per element.
#[test]
fn scalar_functions_lift_over_blocks() {
    let workbook = fixture();
    assert_eq!(
        values("TRIM(A1:A2)", &workbook),
        vec![t("pear"), t("apple")]
    );
    assert_eq!(values("LEFT(A1:A2,2)", &workbook), vec![t("  "), t("ap")]);
    assert_eq!(values("ROUND(B1:B2/3,2)", &workbook), vec![n(1.0), n(0.33)]);
}

#[test]
fn operators_broadcast_elementwise() {
    let workbook = fixture();
    assert_eq!(
        values("B1:B4>2", &workbook),
        vec![
            CellValue::Bool { value: true },
            CellValue::Bool { value: false },
            CellValue::Bool { value: true },
            CellValue::Bool { value: false },
        ]
    );
    assert_eq!(values("-B1:B2", &workbook), vec![n(-3.0), n(-1.0)]);
    assert_eq!(
        arrayed("B1:B2*_xlfn.SEQUENCE(1,2)", &workbook),
        (2, 2, vec![n(3.0), n(6.0), n(1.0), n(2.0)])
    );
}

/// the ctrl-shift-enter idiom: an aggregate consuming a computed block.
#[test]
fn aggregates_consume_computed_blocks() {
    let workbook = fixture();
    assert_eq!(
        values("SUM(IF(A1:A4=\"apple\",B1:B4,0))", &workbook),
        vec![n(3.0)]
    );
    assert_eq!(
        values("MAX(IF(A1:A4=\"apple\",B1:B4,0))", &workbook),
        vec![n(2.0)]
    );
    assert_eq!(values("MATCH(4,B1:B4*1,0)", &workbook), vec![n(3.0)]);
    assert_eq!(
        values("SUMPRODUCT(--(A1:A4=\"apple\"),B1:B4)", &workbook),
        vec![n(3.0)]
    );
}

#[test]
fn iferror_replaces_only_the_failing_elements() {
    let workbook = fixture();
    assert_eq!(
        values("IFERROR(B1:B4/{1;0;1;0},\"x\")", &workbook),
        vec![n(3.0), t("x"), n(4.0), t("x")]
    );
}

/// a bare range is still `#VALUE!` in the scalar evaluator.
#[test]
fn scalar_evaluation_is_unchanged() {
    let workbook = fixture();
    let expression = parse_formula("B1:B4").unwrap();
    let context = EvalContext::new(&workbook, SheetId(0));
    assert_eq!(
        xlsx_calc::evaluate(&expression, &context),
        CellValue::Error {
            value: ErrorValue::Value
        }
    );
}

fn spilling_workbook(formula: &str, reference: &str) -> (Workbook, SheetId) {
    let mut workbook = fixture();
    let sheet = workbook.sheet_mut(SheetId(0)).unwrap();
    sheet.set_cell(
        a1(reference.split(':').next().unwrap()),
        Cell {
            value: CellValue::Empty,
            formula: Some(formula.to_string()),
            style: None,
        },
    );
    sheet.set_array_formula(a1(reference.split(':').next().unwrap()), range(reference));
    (workbook, SheetId(0))
}

#[test]
fn an_array_formula_writes_its_whole_rectangle() {
    let (mut workbook, sheet) = spilling_workbook("_xlfn._xlws.SORT(A1:B4,2)", "D1:E4");
    rebuild_and_recalc_all(&mut workbook, None);
    let value = |address: &str| {
        workbook
            .sheet(sheet)
            .unwrap()
            .cell(a1(address))
            .unwrap()
            .value
            .clone()
    };
    assert_eq!(value("D1"), t("apple"));
    assert_eq!(value("E1"), n(1.0));
    assert_eq!(value("D4"), t("pear"));
    assert_eq!(value("E4"), n(4.0));
    assert_eq!(
        workbook.sheet(sheet).unwrap().array_formula(a1("D1")),
        Some(range("D1:E4"))
    );
}

/// a smaller result retires the cells the previous one reached.
#[test]
fn a_shrinking_result_clears_what_it_no_longer_fills() {
    let (mut workbook, sheet) = spilling_workbook("_xlfn._xlws.FILTER(B1:B4,B1:B4>2)", "D1:D4");
    rebuild_and_recalc_all(&mut workbook, None);
    let sheet = workbook.sheet(sheet).unwrap();
    assert_eq!(sheet.cell(a1("D1")).unwrap().value, n(3.0));
    assert_eq!(sheet.cell(a1("D2")).unwrap().value, n(4.0));
    assert_eq!(sheet.array_formula(a1("D1")), Some(range("D1:D2")));
    for address in ["D3", "D4"] {
        assert!(
            sheet
                .cell(a1(address))
                .is_none_or(|cell| cell.value == CellValue::Empty)
        );
    }
}

/// a single value repeats across the rectangle the file recorded.
#[test]
fn a_single_value_fills_the_recorded_rectangle() {
    let (mut workbook, sheet) = spilling_workbook("SUM(B1:B4)", "D1:E2");
    rebuild_and_recalc_all(&mut workbook, None);
    let sheet = workbook.sheet(sheet).unwrap();
    for address in ["D1", "E1", "D2", "E2"] {
        assert_eq!(sheet.cell(a1(address)).unwrap().value, n(10.0));
    }
}

/// a result growing past its recorded rectangle must not overwrite what the
/// author put there: `#SPILL!` from the anchor, obstruction untouched.
#[test]
fn an_obstructed_rectangle_reports_spill_and_keeps_the_obstruction() {
    for obstruction in [
        Cell {
            value: n(99.0),
            ..Cell::default()
        },
        Cell {
            value: CellValue::Empty,
            formula: Some("1+1".into()),
            style: None,
        },
    ] {
        let (mut workbook, sheet) = spilling_workbook("_xlfn.SEQUENCE(3)", "D1");
        let authored = obstruction.formula.clone();
        workbook
            .sheet_mut(sheet)
            .unwrap()
            .set_cell(a1("D2"), obstruction);
        rebuild_and_recalc_all(&mut workbook, None);
        let sheet = workbook.sheet(sheet).unwrap();
        assert_eq!(
            sheet.cell(a1("D1")).unwrap().value,
            CellValue::Error {
                value: ErrorValue::Spill
            }
        );
        assert_eq!(sheet.cell(a1("D2")).unwrap().formula, authored);
        if authored.is_none() {
            assert_eq!(sheet.cell(a1("D2")).unwrap().value, n(99.0));
        }
        assert!(
            sheet
                .cell(a1("D3"))
                .is_none_or(|cell| cell.value == CellValue::Empty)
        );
        assert_eq!(sheet.array_formula(a1("D1")), Some(range("D1")));
    }
}

/// cells a previous result filled belong to the spill, not to the author, so a
/// workbook opened with its cached spill still recalculates.
#[test]
fn cells_the_previous_result_filled_do_not_block_it() {
    let (mut workbook, sheet) = spilling_workbook("_xlfn.SEQUENCE(3)", "D1:D3");
    for address in ["D2", "D3"] {
        put(workbook.sheet_mut(sheet).unwrap(), address, n(7.0));
    }
    rebuild_and_recalc_all(&mut workbook, None);
    let sheet = workbook.sheet(sheet).unwrap();
    for (address, expected) in [("D1", 1.0), ("D2", 2.0), ("D3", 3.0)] {
        assert_eq!(sheet.cell(a1(address)).unwrap().value, n(expected));
    }
}

/// a formula reading a spilled cell must be evaluated after the spill lands.
#[test]
fn a_reader_of_a_spilled_cell_recalculates_after_it() {
    let (mut workbook, sheet) = spilling_workbook("_xlfn.SEQUENCE(4,1,10,10)", "D1:D4");
    workbook.sheet_mut(sheet).unwrap().set_cell(
        a1("F1"),
        Cell {
            value: CellValue::Empty,
            formula: Some("D4*2".into()),
            style: None,
        },
    );
    rebuild_and_recalc_all(&mut workbook, None);
    assert_eq!(
        workbook.sheet(sheet).unwrap().cell(a1("F1")).unwrap().value,
        n(80.0)
    );
}

/// a growing rectangle reschedules readers of the cells it newly covers, on the
/// incremental path that never saw them.
#[test]
fn a_growing_spill_reschedules_readers_of_the_cells_it_uncovers() {
    let (mut workbook, sheet) = spilling_workbook("_xlfn._xlws.FILTER(B1:B4,B1:B4>3)", "D1");
    workbook.sheet_mut(sheet).unwrap().set_cell(
        a1("F1"),
        Cell {
            value: CellValue::Empty,
            formula: Some("D2*10".into()),
            style: None,
        },
    );
    let (mut graph, _) = rebuild_and_recalc_all(&mut workbook, None);
    assert_eq!(
        workbook.sheet(sheet).unwrap().cell(a1("F1")).unwrap().value,
        n(0.0)
    );

    put(workbook.sheet_mut(sheet).unwrap(), "B2", n(9.0));
    recalc_after(&mut workbook, &mut graph, &[(sheet, a1("B2"))], None);
    let sheet = workbook.sheet(sheet).unwrap();
    assert_eq!(sheet.cell(a1("D1")).unwrap().value, n(9.0));
    assert_eq!(sheet.cell(a1("D2")).unwrap().value, n(4.0));
    assert_eq!(sheet.cell(a1("F1")).unwrap().value, n(40.0));
}

/// an output far larger than both inputs is refused before it is reserved.
#[test]
fn an_oversized_product_is_refused_before_allocation() {
    let workbook = fixture();
    for formula in [
        "MMULT(_xlfn.SEQUENCE(550000,1),_xlfn.SEQUENCE(1,550000))",
        "_xlfn.HSTACK(_xlfn.SEQUENCE(600000,1),_xlfn.SEQUENCE(600000,1))",
        "_xlfn.VSTACK(_xlfn.SEQUENCE(1,600000),_xlfn.SEQUENCE(1,600000))",
        "_xlfn.EXPAND(B1:B2,900000,900000)",
    ] {
        assert_eq!(
            values(formula, &workbook),
            vec![CellValue::Error {
                value: ErrorValue::Num
            }],
            "{formula}"
        );
    }
}

/// a generated block past the limit reports `#NUM!` rather than allocating.
#[test]
fn oversized_results_are_refused() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn.SEQUENCE(1048576,16384)", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Num
        }]
    );
    const { assert!(MAX_SPILL_CELLS <= 1_048_576) };
}

/// whole-column references materialize only the rows the sheet uses.
#[test]
fn whole_column_blocks_stop_at_the_used_range() {
    let workbook = fixture();
    assert_eq!(arrayed("A:A", &workbook).0, 4);
    assert_eq!(values("_xlfn.TOCOL(B:B,1)", &workbook).len(), 4);
}
