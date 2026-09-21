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

/// `LET` binds each name for every later value and for the calculation, and a
/// later binding of the same name shadows the earlier one.
#[test]
fn let_binds_names_in_order() {
    let workbook = fixture();
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x*3,_xlpm.x+_xlpm.y)",
            &workbook
        ),
        vec![n(8.0)]
    );
    assert_eq!(
        values("_xlfn.LET(_xlpm.x,2,_xlpm.x,5,_xlpm.x)", &workbook),
        vec![n(5.0)]
    );
    assert_eq!(
        arrayed("_xlfn.LET(_xlpm.b,B1:B4,_xlpm.b*2)", &workbook).2,
        vec![n(6.0), n(2.0), n(8.0), n(4.0)]
    );
}

/// a name is unbound again once its `LET` returns, and an even argument count
/// leaves the calculation missing.
#[test]
fn let_refuses_malformed_and_unbound_uses() {
    let workbook = fixture();
    for formula in [
        "_xlfn.LET(_xlpm.x,1)",
        "_xlfn.LET(_xlpm.x,1,2,_xlpm.x)",
        "_xlfn.LET(1,1,1)",
    ] {
        assert!(
            matches!(values(formula, &workbook)[..], [CellValue::Error { .. }]),
            "{formula}"
        );
    }
    assert_eq!(
        values("_xlfn.LET(_xlpm.x,1,_xlpm.x)+_xlpm.x", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Name
        }]
    );
}

/// a name bound to a plain reference still reaches a callee that wants an
/// area, so `ROWS` counts rows rather than measuring a block.
#[test]
fn a_name_bound_to_a_range_stays_a_reference() {
    let workbook = fixture();
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.r,A1:A4,ROWS(_xlpm.r)+COLUMNS(_xlpm.r))",
            &workbook
        ),
        vec![n(5.0)]
    );
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.r,B1:B4,SUM(OFFSET(_xlpm.r,1,0,2,1)))",
            &workbook
        ),
        vec![n(5.0)]
    );
}

/// the callback builtins each call a `LAMBDA` and keep their own output shape.
#[test]
fn callbacks_invoke_lambdas() {
    let workbook = fixture();
    assert_eq!(
        arrayed(
            "_xlfn.BYROW(A1:B4,_xlfn.LAMBDA(_xlpm.r,COUNTA(_xlpm.r)))",
            &workbook
        ),
        (4, 1, vec![n(2.0), n(2.0), n(2.0), n(2.0)])
    );
    assert_eq!(
        arrayed(
            "_xlfn.BYCOL(B1:B4,_xlfn.LAMBDA(_xlpm.c,SUM(_xlpm.c)))",
            &workbook
        ),
        (1, 1, vec![n(10.0)])
    );
    assert_eq!(
        values(
            "_xlfn.MAP(B1:B4,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*10))",
            &workbook
        ),
        vec![n(30.0), n(10.0), n(40.0), n(20.0)]
    );
    assert_eq!(
        values(
            "_xlfn.REDUCE(0,B1:B4,_xlfn.LAMBDA(_xlpm.a,_xlpm.v,_xlpm.a+_xlpm.v))",
            &workbook
        ),
        vec![n(10.0)]
    );
    assert_eq!(
        values(
            "_xlfn.SCAN(0,B1:B4,_xlfn.LAMBDA(_xlpm.a,_xlpm.v,_xlpm.a+_xlpm.v))",
            &workbook
        ),
        vec![n(3.0), n(4.0), n(8.0), n(10.0)]
    );
    assert_eq!(
        arrayed(
            "_xlfn.MAKEARRAY(2,3,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*10+_xlpm.c))",
            &workbook
        ),
        (
            2,
            3,
            vec![n(11.0), n(12.0), n(13.0), n(21.0), n(22.0), n(23.0)]
        )
    );
}

/// a lambda parameter shadows an outer `LET` name for the call only, and a
/// free name still resolves through the caller's bindings.
#[test]
fn lambda_parameters_shadow_and_close_over() {
    let workbook = fixture();
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.k,100,_xlfn.MAP(B1:B2,_xlfn.LAMBDA(_xlpm.v,_xlpm.v+_xlpm.k)))",
            &workbook
        ),
        vec![n(103.0), n(101.0)]
    );
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.v,7,_xlfn.MAP(B1:B1,_xlfn.LAMBDA(_xlpm.v,_xlpm.v)))",
            &workbook
        ),
        vec![n(3.0)]
    );
}

/// `BYROW` hands the lambda the row it is called with as a reference, so a
/// callee that needs one still works.
#[test]
fn byrow_passes_each_row_as_a_reference() {
    let workbook = fixture();
    assert_eq!(
        values(
            "_xlfn.BYROW(B2:B3,_xlfn.LAMBDA(_xlpm.r,SUM(OFFSET(_xlpm.r,0,-1,1,2))))",
            &workbook
        ),
        vec![n(1.0), n(4.0)]
    );
}

/// a callback argument that is not a lambda, and a lambda given the wrong
/// number of arguments, both report `#VALUE!`.
#[test]
fn callbacks_refuse_a_non_lambda() {
    let workbook = fixture();
    for formula in ["_xlfn.MAP(B1:B4,1)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)"] {
        assert_eq!(
            values(formula, &workbook),
            vec![CellValue::Error {
                value: ErrorValue::Value
            }],
            "{formula}"
        );
    }
    assert_eq!(
        values(
            "_xlfn.BYROW(B1:B4,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,1))",
            &workbook
        ),
        vec![
            CellValue::Error {
                value: ErrorValue::Value
            };
            4
        ]
    );
}

/// `TEXTSPLIT` splits rows first, then columns, padding short rows.
#[test]
fn textsplit_builds_a_grid() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn.TEXTSPLIT(\"a,b;c\",\",\",\";\")", &workbook),
        (
            2,
            2,
            vec![
                t("a"),
                t("b"),
                t("c"),
                CellValue::Error {
                    value: ErrorValue::NA
                }
            ]
        )
    );
    assert_eq!(
        arrayed(
            "_xlfn.TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,\"-\")",
            &workbook
        ),
        (2, 2, vec![t("a"), t("b"), t("c"), t("-")])
    );
    assert_eq!(
        values("_xlfn.TEXTSPLIT(\"a,,b\",\",\",,TRUE)", &workbook),
        vec![t("a"), t("b")]
    );
    assert_eq!(
        values("_xlfn.TEXTSPLIT(\"aXbxc\",\"x\",,,1)", &workbook),
        vec![t("a"), t("b"), t("c")]
    );
    assert_eq!(
        values("_xlfn.TEXTSPLIT(\"a+-b\",{\"+\",\"+-\"})", &workbook),
        vec![t("a"), t("b")]
    );
    assert_eq!(
        values("_xlfn.TEXTSPLIT(\"abc\",\"\")", &workbook),
        vec![t("abc")]
    );
}

/// `TEXTJOIN`/`CONCAT` read every cell of a computed block, not just its
/// top-left value.
#[test]
fn joins_read_whole_blocks() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn.TEXTJOIN(\",\",TRUE,B1:B4*2)", &workbook),
        vec![t("6,2,8,4")]
    );
    assert_eq!(
        values("_xlfn.CONCAT(_xlfn.SEQUENCE(3))", &workbook),
        vec![t("123")]
    );
}

/// a criterion or lookup key given as a block answers once per element while
/// the range arguments still arrive as references.
#[test]
fn criteria_and_lookup_keys_lift_elementwise() {
    let workbook = fixture();
    assert_eq!(
        values("COUNTIF(A1:A4,_xlfn.UNIQUE(A1:A4))", &workbook),
        vec![n(1.0), n(2.0), n(1.0)]
    );
    assert_eq!(
        values("SUMIF(A1:A4,_xlfn.UNIQUE(A1:A4),B1:B4)", &workbook),
        vec![n(3.0), n(3.0), n(4.0)]
    );
}

/// `INDEX` with a zero or omitted index answers with the whole row or column.
#[test]
fn index_returns_a_whole_axis() {
    let workbook = fixture();
    assert_eq!(
        arrayed("INDEX(A1:B4,,2)", &workbook),
        (4, 1, vec![n(3.0), n(1.0), n(4.0), n(2.0)])
    );
    assert_eq!(
        arrayed("INDEX(A1:B4,2,0)", &workbook),
        (1, 2, vec![t("apple"), n(1.0)])
    );
}

/// `INDEX(block, {3;1})` answers once per index, over the rectangle the row
/// and column indices broadcast to.
#[test]
fn index_answers_once_per_array_index() {
    let workbook = fixture();
    assert_eq!(
        arrayed("INDEX(A1:A4,_xlfn.SEQUENCE(3,,4,-1))", &workbook),
        (3, 1, vec![t("apple"), t("pear"), t("apple")])
    );
    assert_eq!(
        arrayed("INDEX(A1:B4,{2;3},{1,2})", &workbook),
        (2, 2, vec![t("apple"), n(1.0), t("pear"), n(4.0)])
    );
    assert_eq!(
        arrayed("INDEX(A1:A4,{1;9})", &workbook),
        (
            2,
            1,
            vec![
                t("  pear "),
                CellValue::Error {
                    value: ErrorValue::Ref
                }
            ]
        )
    );
}

/// `WRAPROWS`/`WRAPCOLS` cut a vector into a rectangle, padding the tail.
#[test]
fn wrapping_a_vector_pads_its_last_group() {
    let workbook = fixture();
    assert_eq!(
        arrayed("_xlfn.WRAPROWS(_xlfn.SEQUENCE(5),2,0)", &workbook),
        (3, 2, vec![n(1.0), n(2.0), n(3.0), n(4.0), n(5.0), n(0.0)])
    );
    assert_eq!(
        arrayed("_xlfn.WRAPCOLS(_xlfn.SEQUENCE(5),2)", &workbook),
        (
            2,
            3,
            vec![
                n(1.0),
                n(3.0),
                n(5.0),
                n(2.0),
                n(4.0),
                CellValue::Error {
                    value: ErrorValue::NA
                }
            ]
        )
    );
}

/// a whole-column criteria range costs the rows the sheet uses, so one
/// `COUNTIFS` per key of a spilled block stays inside the evaluation budget.
#[test]
fn whole_column_criteria_cost_the_used_rows() {
    let workbook = fixture();
    assert_eq!(
        values(r#"COUNTIFS(A:A,_xlfn.UNIQUE(A1:A4),B:B,">1")"#, &workbook),
        vec![n(1.0), n(1.0), n(1.0)]
    );
    assert_eq!(
        values(r#"SUMIFS(B:B,A:A,_xlfn.UNIQUE(A1:A4))"#, &workbook),
        vec![n(3.0), n(3.0), n(4.0)]
    );
}

/// `XLOOKUP` over a two-dimensional result answers with the whole row its
/// lookup vector picks, and accepts a vector the formula computed.
#[test]
fn xlookup_picks_a_row_of_a_block() {
    let workbook = fixture();
    assert_eq!(
        arrayed(r#"_xlfn.XLOOKUP("pear",A1:A4,A1:B4)"#, &workbook),
        (1, 2, vec![t("pear"), n(4.0)])
    );
    assert_eq!(
        values(r#"_xlfn.XLOOKUP(1,(A1:A4="apple")*1,B1:B4)"#, &workbook),
        vec![n(1.0)]
    );
    assert_eq!(
        values(r#"_xlfn.XLOOKUP("fig",A1:A4,B1:B4,"none")"#, &workbook),
        vec![t("none")]
    );
}

/// match mode picks the nearest value either side of the key; search mode -1
/// reads the vector from the end.
#[test]
fn xlookup_honours_match_and_search_modes() {
    let workbook = fixture();
    assert_eq!(
        values("_xlfn.XLOOKUP(3.5,B1:B4,A1:A4,,-1)", &workbook),
        vec![t("  pear ")]
    );
    assert_eq!(
        values("_xlfn.XLOOKUP(3.5,B1:B4,A1:A4,,1)", &workbook),
        vec![t("pear")]
    );
    assert_eq!(
        values(r#"_xlfn.XMATCH("apple",A1:A4)"#, &workbook),
        vec![n(2.0)]
    );
    assert_eq!(
        values(r#"_xlfn.XMATCH("apple",A1:A4,0,-1)"#, &workbook),
        vec![n(4.0)]
    );
    assert_eq!(
        values(r#"_xlfn.XMATCH({"pear";"apple"},A1:A4)"#, &workbook),
        vec![n(3.0), n(2.0)]
    );
}

/// a text builtin with no array-aware form still answers once per element.
#[test]
fn text_builtins_lift_over_a_range() {
    let workbook = fixture();
    assert_eq!(
        values(r#"_xlfn.TEXTBEFORE(A1:A4,"p")"#, &workbook),
        vec![t("  "), t("a"), t(""), t("a")]
    );
    assert_eq!(
        values("ISODD(B1:B4)", &workbook),
        vec![
            CellValue::Bool { value: true },
            CellValue::Bool { value: true },
            CellValue::Bool { value: false },
            CellValue::Bool { value: false },
        ]
    );
}

/// a callback body that reads a bound block pays for each read, so a large one
/// cannot be copied for free once per output cell.
#[test]
fn reading_a_bound_block_charges_the_budget() {
    let workbook = fixture();
    let body = "_xlfn.MAKEARRAY(20,20,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,ROWS(_xlpm.x)))";
    let large = values(
        &format!("_xlfn.LET(_xlpm.x,_xlfn.SEQUENCE(60000),{body})"),
        &workbook,
    );
    assert_eq!(large.len(), 400);
    assert!(large.iter().any(|value| matches!(
        value,
        CellValue::Error {
            value: ErrorValue::Num
        }
    )));
    let small = values(
        &format!("_xlfn.LET(_xlpm.x,_xlfn.SEQUENCE(4),{body})"),
        &workbook,
    );
    assert!(small.iter().all(|value| *value == n(4.0)));
}

/// `LINEST` lays its coefficients out right to left with the intercept last,
/// so a two-predictor fit of `y = 2*x1 + 3*x2 + 5` answers `[3, 2, 5]`.
#[test]
fn linest_returns_coefficients_in_excel_order() {
    let workbook = fixture();
    let ys = "{13;12;23;22;33;32;43}";
    let xs = "{1,2;2,1;3,4;4,3;5,6;6,5;7,8}";
    let (rows, cols, values) = arrayed(&format!("LINEST({ys},{xs})"), &workbook);
    assert_eq!((rows, cols), (1, 3));
    for (value, want) in values.iter().zip([3.0, 2.0, 5.0]) {
        assert!(matches!(value, CellValue::Number { value } if (value - want).abs() < 1e-9));
    }
}

/// a single predictor with `stats` on reports five rows, and R² for an exact
/// fit is 1 with a residual sum of squares of 0.
#[test]
fn linest_reports_the_statistics_block() {
    let workbook = fixture();
    let (rows, cols, values) = arrayed("LINEST({10;20;30;40;50},{1;2;3;4;5},TRUE,TRUE)", &workbook);
    assert_eq!((rows, cols), (5, 2));
    let close = |index: usize, want: f64| {
        assert!(
            matches!(&values[index], CellValue::Number { value } if (value - want).abs() < 1e-9),
            "cell {index} was {:?}, wanted {want}",
            values[index]
        );
    };
    close(0, 10.0);
    close(1, 0.0);
    close(4, 1.0);
    close(7, 3.0);
    close(9, 0.0);
}

/// fewer observations than terms has no unique fit.
#[test]
fn linest_refuses_an_underdetermined_fit() {
    let workbook = fixture();
    assert_eq!(
        values("LINEST({1;2},{1,2;3,4})", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Num
        }]
    );
}

/// `RANK` over a range of values answers one rank per value, the shape the
/// classic `LOOKUP(2, 1/(RANK(..)=n), ..)` idiom depends on.
#[test]
fn rank_ranks_every_value_it_is_given() {
    let workbook = fixture();
    assert_eq!(
        arrayed("RANK(B1:B4,B1:B4,1)", &workbook),
        (4, 1, vec![n(3.0), n(1.0), n(4.0), n(2.0)])
    );
    assert_eq!(
        arrayed("RANK(B1:B4,B1:B4)", &workbook),
        (4, 1, vec![n(2.0), n(4.0), n(1.0), n(3.0)])
    );
}

/// `TREND` evaluates the line `LINEST` fits: a perfect `y = 10x` fit predicts
/// exactly, in the shape of `new_x`, and falls back to the known predictors.
#[test]
fn trend_predicts_along_the_fitted_line() {
    let workbook = fixture();
    let close = |formula: &str, want: &[f64]| {
        let got = values(formula, &workbook);
        assert_eq!(got.len(), want.len(), "{formula}: {got:?}");
        for (value, want) in got.iter().zip(want) {
            assert!(
                matches!(value, CellValue::Number { value } if (value - want).abs() < 1e-9),
                "{formula}: {value:?} != {want}"
            );
        }
    };
    close("TREND({10;20;30;40;50},{1;2;3;4;5},{6;7})", &[60.0, 70.0]);
    close("TREND({10,20,30,40,50},{1,2,3,4,5},6)", &[60.0]);
    close(
        "TREND({10;20;30;40;50},{1;2;3;4;5})",
        &[10.0, 20.0, 30.0, 40.0, 50.0],
    );
    close("TREND({10;20;30;40;50})", &[10.0, 20.0, 30.0, 40.0, 50.0]);
    close(
        "TREND({13;12;23;22;33;32;43},{1,2;2,1;3,4;4,3;5,6;6,5;7,8},{1,2;2,1})",
        &[13.0, 12.0],
    );
    assert_eq!(
        arrayed("TREND({10;20;30;40;50},{1;2;3;4;5},{6;7})", &workbook).0,
        2
    );
    assert_eq!(
        values("TREND({1;2},{1,2;3,4})", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Num
        }]
    );
    assert_eq!(
        values("TREND({10;20;30},{1;2})", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Ref
        }]
    );
    assert_eq!(
        values("TREND()", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Value
        }]
    );
}

/// `FREQUENCY` answers a column one taller than its bins, counts each value
/// in the first interval that holds it, and ignores everything non-numeric.
#[test]
fn frequency_bins_values_into_one_column() {
    let workbook = fixture();
    assert_eq!(
        arrayed("FREQUENCY({1;2;3;4;5},{2;4})", &workbook),
        (3, 1, vec![n(2.0), n(2.0), n(1.0)])
    );
    assert_eq!(
        arrayed("FREQUENCY({1;2;3;4;5},{4;2})", &workbook),
        (3, 1, vec![n(2.0), n(2.0), n(1.0)])
    );
    assert_eq!(
        arrayed("FREQUENCY({1,\"x\",TRUE,3},{2})", &workbook),
        (2, 1, vec![n(1.0), n(1.0)])
    );
    assert_eq!(
        arrayed("FREQUENCY({1;2;3},{3;3})", &workbook),
        (3, 1, vec![n(3.0), n(0.0), n(0.0)])
    );
    assert_eq!(
        values("FREQUENCY({1;2;3},{9})", &workbook),
        vec![n(3.0), n(0.0)]
    );
    assert_eq!(
        values("FREQUENCY({1;2;3},{1/0})", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Div0
        }]
    );
    assert_eq!(
        values("FREQUENCY({1;2;3})", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Value
        }]
    );
}

/// the longest-run idiom: `FREQUENCY` over the row numbers of the matching
/// rows, binned by the row numbers of the rest, and the widest gap wins.
#[test]
fn frequency_measures_the_longest_run() {
    let workbook = fixture();
    assert_eq!(
        values(
            "MAX(FREQUENCY(IF(A1:A4=\"apple\",ROW(A1:A4)),IF(A1:A4<>\"apple\",ROW(A1:A4))))",
            &workbook
        ),
        vec![n(1.0)]
    );
}
