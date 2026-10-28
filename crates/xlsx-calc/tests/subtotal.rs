//! SUBTOTAL/AGGREGATE ignore cells in their references whose own formula
//! contains a SUBTOTAL or AGGREGATE call, so a grand total over group
//! subtotals does not double count.

use xlsx_calc::{EvalContext, evaluate, parse_formula, rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellRef, CellValue, DefinedName, ErrorValue, Sheet, SheetId, Workbook};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn a1(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn plain(value: CellValue) -> Cell {
    Cell {
        value,
        ..Cell::default()
    }
}

/// a formula cell carrying the cached value excel would have stored for it.
fn formula(cached: CellValue, src: &str) -> Cell {
    Cell {
        value: cached,
        formula: Some(src.into()),
        ..Cell::default()
    }
}

fn workbook(sheet: &str, cells: &[(&str, Cell)]) -> Workbook {
    let mut wb = Workbook::default();
    let mut s = Sheet::new(sheet);
    for (address, cell) in cells {
        s.set_cell(a1(address), cell.clone());
    }
    wb.sheets.push(s);
    wb
}

fn eval_in(wb: &Workbook, src: &str) -> CellValue {
    evaluate(
        &parse_formula(src).expect("parse"),
        &EvalContext::new(wb, SheetId(0)),
    )
}

fn check_in(wb: &Workbook, cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval_in(wb, src), *want, "formula {src:?}");
    }
}

/// the workbook from the report: detail rows summing to 1,608, with a group
/// subtotal after each group also summing to 1,608. a grand total that reads
/// every cell therefore reports 3,216 instead of 1,608.
fn grouped() -> Workbook {
    workbook(
        "Sheet1",
        &[
            ("A1", plain(n(500.0))),
            ("A2", plain(n(300.0))),
            ("A3", plain(n(200.0))),
            ("A4", formula(n(1000.0), "SUBTOTAL(9,A1:A3)")),
            ("A5", plain(n(100.0))),
            ("A6", plain(n(200.0))),
            ("A7", plain(n(8.0))),
            ("A8", formula(n(308.0), "SUBTOTAL(9,A5:A7)")),
            ("A9", plain(n(300.0))),
            ("A10", formula(n(300.0), "SUBTOTAL(9,A9:A9)")),
        ],
    )
}

#[test]
fn a_grand_subtotal_skips_the_group_subtotals_it_covers() {
    let wb = grouped();
    let detail_total = 500.0 + 300.0 + 200.0 + 100.0 + 200.0 + 8.0 + 300.0;
    assert_eq!(detail_total, 1608.0);
    check_in(
        &wb,
        &[
            ("SUBTOTAL(9,A1:A10)", n(detail_total)),
            ("SUBTOTAL(109,A1:A10)", n(detail_total)),
            ("SUBTOTAL(9,A1:A4)", n(1000.0)),
            // a single nested cell read through a 1x1 reference is skipped too
            ("SUBTOTAL(9,A4)", n(0.0)),
            // literal and computed arguments are not cells, so never skipped
            ("SUBTOTAL(9,A1:A3,5)", n(1005.0)),
            ("SUBTOTAL(9,{10,20})", n(30.0)),
        ],
    );
}

#[test]
fn subtotal_109_and_101_average_skip_nested_cells() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            ("A2", plain(n(20.0))),
            ("A3", plain(n(30.0))),
            ("A4", formula(n(60.0), "SUBTOTAL(9,A1:A3)")),
        ],
    );
    check_in(
        &wb,
        &[
            ("SUBTOTAL(109,A1:A4)", n(60.0)),
            ("SUBTOTAL(1,A1:A4)", n(20.0)),
            ("SUBTOTAL(101,A1:A4)", n(20.0)),
        ],
    );
}

#[test]
fn count_and_counta_skip_nested_cells() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            (
                "A2",
                plain(CellValue::Text {
                    value: "label".into(),
                }),
            ),
            ("A3", plain(n(30.0))),
            ("A4", formula(n(40.0), "SUBTOTAL(9,A1:A3)")),
        ],
    );
    check_in(
        &wb,
        &[
            // COUNT sees only the two detail numbers, not the nested result
            ("SUBTOTAL(2,A1:A4)", n(2.0)),
            ("SUBTOTAL(102,A1:A4)", n(2.0)),
            // COUNTA sees the details including the label, not the nested cell
            ("SUBTOTAL(3,A1:A4)", n(3.0)),
            ("SUBTOTAL(103,A1:A4)", n(3.0)),
        ],
    );
}

#[test]
fn a_subtotal_call_anywhere_in_a_formula_marks_the_cell_nested() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            ("A2", plain(n(20.0))),
            // SUBTOTAL under an operator still marks the cell
            ("A3", formula(n(60.0), "SUBTOTAL(9,A1:A2)*2")),
            // and so does SUBTOTAL in a branch, and a nested AGGREGATE
            ("A4", formula(n(30.0), "IF(TRUE,SUBTOTAL(9,A1:A2),0)")),
            ("A5", formula(n(30.0), "AGGREGATE(9,4,A1:A2)")),
            // a plain SUM formula is an ordinary computed cell: it counts
            ("A6", formula(n(30.0), "SUM(A1:A2)")),
        ],
    );
    check_in(
        &wb,
        &[
            ("SUBTOTAL(9,A1:A6)", n(60.0)),
            ("AGGREGATE(9,0,A1:A6)", n(60.0)),
        ],
    );
}

#[test]
fn a_cross_sheet_range_skips_nested_cells_on_that_sheet() {
    let mut wb = workbook("Sheet1", &[]);
    let mut data = Sheet::new("Data");
    for (address, cell) in [
        ("A1", plain(n(5.0))),
        ("A2", plain(n(7.0))),
        ("A3", formula(n(12.0), "SUBTOTAL(9,A1:A2)")),
        ("A4", plain(n(8.0))),
    ] {
        data.set_cell(a1(address), cell);
    }
    wb.sheets.push(data);
    check_in(
        &wb,
        &[
            ("SUBTOTAL(9,Data!A1:A4)", n(20.0)),
            ("AGGREGATE(9,0,Data!A1:A4)", n(20.0)),
            ("AGGREGATE(9,4,Data!A1:A4)", n(32.0)),
        ],
    );
}

#[test]
fn aggregate_options_decide_whether_nested_cells_count() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            ("A2", plain(n(20.0))),
            ("A3", formula(n(30.0), "SUBTOTAL(9,A1:A2)")),
            (
                "A4",
                formula(
                    CellValue::Error {
                        value: ErrorValue::Div0,
                    },
                    "1/0",
                ),
            ),
        ],
    );
    check_in(
        &wb,
        &[
            // options 0-3 ignore nested subtotals
            ("AGGREGATE(9,0,A1:A3)", n(30.0)),
            ("AGGREGATE(9,1,A1:A3)", n(30.0)),
            // an omitted options argument reads as 0
            ("AGGREGATE(9,,A1:A3)", n(30.0)),
            // options 4-7 include them
            ("AGGREGATE(9,4,A1:A3)", n(60.0)),
            ("AGGREGATE(9,5,A1:A3)", n(60.0)),
            // option 6 includes nested cells and ignores the error cell
            ("AGGREGATE(9,6,A1:A4)", n(60.0)),
            ("AGGREGATE(9,7,A1:A4)", n(60.0)),
            // options 2 and 3 ignore nested cells and the error cell
            ("AGGREGATE(9,2,A1:A4)", n(30.0)),
            ("AGGREGATE(9,3,A1:A4)", n(30.0)),
        ],
    );
}

#[test]
fn aggregate_large_over_a_reference_applies_the_same_rule() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            ("A2", plain(n(20.0))),
            ("A3", formula(n(30.0), "SUBTOTAL(9,A1:A2)")),
            ("A4", plain(n(5.0))),
        ],
    );
    check_in(
        &wb,
        &[
            // reference form, option 6: the nested subtotal is the largest cell
            ("AGGREGATE(14,6,A1:A4,1)", n(30.0)),
            // reference form, option 0: it is skipped, so 20 is largest
            ("AGGREGATE(14,0,A1:A4,1)", n(20.0)),
            ("AGGREGATE(15,0,A1:A4,1)", n(5.0)),
            // array form: a computed array has no cells to skip
            ("AGGREGATE(14,6,{10,30,20},1)", n(30.0)),
            ("AGGREGATE(14,0,A1:A4*1,1)", n(30.0)),
        ],
    );
}

#[test]
fn unions_and_defined_names_that_resolve_to_references_skip_nested_cells() {
    let mut wb = grouped();
    wb.defined_names.push(DefinedName {
        name: "Block".into(),
        formula: "Sheet1!$A$1:$A$10".into(),
        local_sheet: None,
        hidden: false,
    });
    check_in(
        &wb,
        &[
            ("SUBTOTAL(9,(A1:A4,A5:A8))", n(1308.0)),
            ("SUBTOTAL(9,Block)", n(1608.0)),
            ("AGGREGATE(9,0,Block)", n(1608.0)),
        ],
    );
}

#[test]
fn editing_a_nested_subtotals_input_updates_the_grand_total() {
    let mut wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(500.0))),
            ("A2", plain(n(300.0))),
            ("A3", plain(n(200.0))),
            ("A4", formula(CellValue::Empty, "SUBTOTAL(9,A1:A3)")),
            ("A5", plain(n(100.0))),
            ("A6", plain(n(200.0))),
            ("A7", plain(n(8.0))),
            ("A8", formula(CellValue::Empty, "SUBTOTAL(9,A5:A7)")),
            ("A9", plain(n(300.0))),
            ("A10", formula(CellValue::Empty, "SUBTOTAL(9,A9:A9)")),
            ("A11", formula(CellValue::Empty, "SUBTOTAL(9,A1:A10)")),
        ],
    );
    let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
    let value = |wb: &Workbook, address: &str| {
        wb.sheet(SheetId(0))
            .unwrap()
            .cell(a1(address))
            .unwrap()
            .value
            .clone()
    };
    assert_eq!(value(&wb, "A4"), n(1000.0));
    assert_eq!(value(&wb, "A11"), n(1608.0));

    wb.sheet_mut(SheetId(0))
        .unwrap()
        .set_cell(a1("A1"), plain(n(600.0)));
    recalc_after(&mut wb, &mut graph, &[(SheetId(0), a1("A1"))], None);
    assert_eq!(value(&wb, "A4"), n(1100.0));
    assert_eq!(value(&wb, "A11"), n(1708.0));
}

#[test]
fn aggregate_ranked_codes_need_exactly_an_array_and_a_k() {
    let wb = workbook(
        "Sheet1",
        &[
            ("A1", plain(n(10.0))),
            ("A2", plain(n(20.0))),
            ("B1", plain(n(5.0))),
            ("B2", plain(n(35.0))),
        ],
    );
    check_in(
        &wb,
        &[
            // codes 14-17 take exactly (array, k) after the options, as the
            // standalone LARGE/SMALL/PERCENTILE/QUARTILE do
            ("AGGREGATE(14,0,A1:A2)", e(ErrorValue::Value)),
            ("AGGREGATE(14,0,A1:A2,B1:B2,1)", e(ErrorValue::Value)),
            ("AGGREGATE(15,0,A1:A2,B1:B2,1)", e(ErrorValue::Value)),
            ("AGGREGATE(16,0,A1:A2,B1:B2,0.5)", e(ErrorValue::Value)),
            ("AGGREGATE(17,0,A1:A2,B1:B2,2)", e(ErrorValue::Value)),
            ("AGGREGATE(15,0,A1:A2)", e(ErrorValue::Value)),
            ("AGGREGATE(16,0,A1:A2)", e(ErrorValue::Value)),
            ("AGGREGATE(17,0,A1:A2)", e(ErrorValue::Value)),
            // with options 6 the wrong arity still reaches the error-ignoring
            // retry, exactly as on main: no data argument is left for the
            // first shape, so it is #DIV/0!, and the second shape ranks the
            // two ranges it was given, largest first
            ("AGGREGATE(14,6,A1:A2)", e(ErrorValue::Div0)),
            ("AGGREGATE(14,6,A1:A2,B1:B2,1)", n(35.0)),
        ],
    );
}
