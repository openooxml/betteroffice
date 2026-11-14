use xlsx_calc::graph::DepGraph;
use xlsx_calc::{derive_unevaluated, rebuild_and_recalc_all, recalc_after};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, DefinedName, ErrorValue, Sheet, SheetId, Workbook,
};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

const NAME: CellValue = CellValue::Error {
    value: ErrorValue::Name,
};

/// a sheet of `(a1, formula, value)` cells; a formula with an empty value is
/// one the file stored no result for.
fn book(cells: &[(&str, Option<&str>, CellValue)]) -> Workbook {
    let mut sheet = Sheet::new("Sheet1");
    for (at, formula, value) in cells {
        sheet.set_cell(
            CellRef::parse_a1(at).unwrap(),
            Cell {
                value: value.clone(),
                formula: formula.map(str::to_owned),
                style: None,
            },
        );
    }
    Workbook {
        sheets: vec![sheet],
        ..Workbook::default()
    }
}

fn name(book: &mut Workbook, name: &str, formula: &str) {
    book.defined_names.push(DefinedName {
        name: name.into(),
        formula: formula.into(),
        local_sheet: None,
        hidden: false,
    });
}

fn value(book: &Workbook, at: &str) -> CellValue {
    book.sheets[0]
        .cell(CellRef::parse_a1(at).unwrap())
        .map_or(CellValue::Empty, |cell| cell.value.clone())
}

fn listed(book: &mut Workbook) -> Vec<(String, Vec<String>)> {
    rebuild_and_recalc_all(book, Some(46_327.5))
        .1
        .unevaluated_cells
        .iter()
        .map(|cell| (cell.cell.to_a1(), cell.functions.clone()))
        .collect()
}

fn entry(at: &str, functions: &[&str]) -> (String, Vec<String>) {
    (
        at.to_owned(),
        functions.iter().map(|name| (*name).to_owned()).collect(),
    )
}

/// the issue's core complaint: a handler's answer to a call the engine does
/// not know replaced what excel saved, with no signal.
#[test]
fn a_handled_unknown_call_keeps_excels_value_and_is_listed() {
    let mut wb = book(&[(
        "A1",
        Some("IFERROR(REPRO_UNKNOWN_FUNCTION(1),0)"),
        number(42.0),
    )]);
    assert_eq!(listed(&mut wb), [entry("A1", &["REPRO_UNKNOWN_FUNCTION"])]);
    assert_eq!(value(&wb, "A1"), number(42.0));
}

#[test]
fn every_handler_leaves_a_saved_value_alone() {
    for formula in [
        "IFNA(NOPE(1),0)",
        "ISERROR(NOPE(1))",
        "IF(ISERROR(NOPE(1)),0,1)",
        "ERROR.TYPE(NOPE(1))",
        "SUM(IFERROR(NOPE(A2:A3),0))",
    ] {
        let mut wb = book(&[("A1", Some(formula), number(42.0))]);
        assert_eq!(listed(&mut wb), [entry("A1", &["NOPE"])], "{formula}");
        assert_eq!(value(&wb, "A1"), number(42.0), "{formula}");
    }
}

#[test]
fn a_bare_unknown_call_is_listed_with_its_name() {
    let mut wb = book(&[("A1", Some("REPRO_UNKNOWN_FUNCTION(1)"), number(999.0))]);
    assert_eq!(listed(&mut wb), [entry("A1", &["REPRO_UNKNOWN_FUNCTION"])]);
    assert_eq!(value(&wb, "A1"), number(999.0));
}

#[test]
fn without_a_saved_value_the_engines_answer_stands_and_is_listed() {
    let mut wb = book(&[
        ("A1", Some("IFERROR(NOPE(1),0)"), CellValue::Empty),
        ("A2", Some("NOPE(1)"), CellValue::Empty),
    ]);
    let reported = [entry("A1", &["NOPE"]), entry("A2", &["NOPE"])];
    assert_eq!(listed(&mut wb), reported);
    assert_eq!(value(&wb, "A1"), number(0.0));
    assert_eq!(value(&wb, "A2"), NAME);
    assert_eq!(listed(&mut wb), reported);
    assert_eq!(value(&wb, "A1"), number(0.0));
    assert_eq!(value(&wb, "A2"), NAME);
}

/// without a saved value there is nothing to keep, so the answer follows its
/// inputs; a saved one stays through them.
#[test]
fn only_the_engines_answer_follows_an_input() {
    let mut wb = book(&[
        ("A1", Some("IFERROR(NOPE(C1),C1)"), CellValue::Empty),
        ("B1", Some("IFERROR(NOPE(C1),C1)"), number(42.0)),
        ("C1", None, number(1.0)),
    ]);
    let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
    assert_eq!(value(&wb, "A1"), number(1.0));
    let input = CellRef::parse_a1("C1").unwrap();
    wb.sheets[0].set_calculated_value(input, number(2.0));
    let result = recalc_after(&mut wb, &mut graph, &[(SheetId(0), input)], None);
    assert_eq!(value(&wb, "A1"), number(2.0));
    assert_eq!(value(&wb, "B1"), number(42.0));
    assert_eq!(result.unevaluated_cells.len(), 2);
}

#[test]
fn readers_of_a_kept_value_keep_theirs_and_are_listed() {
    let mut wb = book(&[
        ("A1", Some("IFERROR(NOPE(1),0)"), number(42.0)),
        ("B1", Some("A1*2"), number(5.0)),
        ("C1", Some("B1+1"), number(6.0)),
        ("D1", Some("1+1"), number(9.0)),
    ]);
    assert_eq!(
        listed(&mut wb),
        [
            entry("A1", &["NOPE"]),
            entry("B1", &["NOPE"]),
            entry("C1", &["NOPE"]),
        ]
    );
    assert_eq!(value(&wb, "B1"), number(5.0));
    assert_eq!(value(&wb, "C1"), number(6.0));
    assert_eq!(value(&wb, "D1"), number(2.0));
}

/// a reader with nothing saved computes from what it reads, the kept value or
/// the engine's answer, and is listed. a reader of an answer writes too: no
/// value of excel's rests on a cell excel saved none for.
#[test]
fn a_reader_computes_when_nothing_of_excels_is_at_stake() {
    let mut wb = book(&[
        ("A1", Some("NOPE(1)"), number(21.0)),
        ("B1", Some("A1*2"), CellValue::Empty),
        ("C1", Some("B1+1"), CellValue::Empty),
        ("A2", Some("NOPE(1)"), CellValue::Empty),
        ("B2", Some("A2*2"), number(5.0)),
    ]);
    assert_eq!(listed(&mut wb).len(), 5);
    assert_eq!(value(&wb, "B1"), number(42.0));
    assert_eq!(value(&wb, "C1"), number(43.0));
    assert_eq!(value(&wb, "B2"), NAME);
}

#[test]
fn an_uncomputable_call_still_holds_the_cell_it_shares_with_an_unknown_one() {
    let mut wb = book(&[
        (
            "A1",
            Some(r#"IFERROR(NOPE(1),WEBSERVICE("http://x"))"#),
            CellValue::Empty,
        ),
        ("B1", Some("A1+1"), CellValue::Empty),
        ("C1", Some(r#"WEBSERVICE("http://x")+NOPE(1)"#), number(7.0)),
    ]);
    assert_eq!(
        listed(&mut wb),
        [
            entry("A1", &["NOPE", "WEBSERVICE"]),
            entry("B1", &["NOPE", "WEBSERVICE"]),
            entry("C1", &["WEBSERVICE"]),
        ]
    );
    assert_eq!(value(&wb, "A1"), CellValue::Empty);
    assert_eq!(value(&wb, "B1"), CellValue::Empty);
    assert_eq!(value(&wb, "C1"), number(7.0));
}

#[test]
fn stored_prefixes_stay_out_of_the_listed_name() {
    for (formula, reported) in [
        ("_xlfn.FUTURE(1)", "FUTURE"),
        ("_xlfn._xlws.FUTURE(1)", "FUTURE"),
        ("IFERROR(_xlfn.FUTURE(1),0)", "FUTURE"),
        ("_xludf.MyFn(1)", "_xludf.MyFn"),
        ("_xll.Quote(1)", "_xll.Quote"),
    ] {
        let mut wb = book(&[("A1", Some(formula), number(42.0))]);
        assert_eq!(listed(&mut wb), [entry("A1", &[reported])], "{formula}");
        assert_eq!(value(&wb, "A1"), number(42.0), "{formula}");
    }
}

/// calling a defined name is a call the engine cannot make, whatever the name
/// holds: it has no way to invoke a lambda, and excel's answer for a name that
/// holds a value is not the engine's to guess.
#[test]
fn a_defined_name_called_as_a_function_keeps_excels_value() {
    for (at, formula) in [
        ("A1", "MyLambda(2)"),
        ("A2", "IFERROR(MyLambda(2),0)"),
        ("A3", "Levy(2)"),
        ("A4", "IFERROR(Levy(2),0)"),
    ] {
        let mut wb = book(&[("A1", Some(formula), number(42.0))]);
        name(&mut wb, "MyLambda", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)");
        name(&mut wb, "Levy", "0.5");
        let called = formula.trim_start_matches("IFERROR(");
        let called = called.split('(').next().unwrap();
        assert_eq!(listed(&mut wb), [entry("A1", &[called])], "{at}");
        assert_eq!(value(&wb, "A1"), number(42.0), "{at}");
    }
}

#[test]
fn a_let_bound_lambda_called_by_name_keeps_excels_value() {
    let mut wb = book(&[(
        "A1",
        Some("_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(2))"),
        number(3.0),
    )]);
    assert_eq!(listed(&mut wb), [entry("A1", &["_xlpm.f"])]);
    assert_eq!(value(&wb, "A1"), number(3.0));
}

#[test]
fn a_call_inside_a_defined_name_reaches_the_cell_using_it() {
    let mut wb = book(&[
        ("A1", Some("Rate*2"), number(42.0)),
        ("A2", Some("IFERROR(Rate,0)"), number(42.0)),
    ]);
    name(&mut wb, "Rate", "NOPE(1)");
    assert_eq!(
        listed(&mut wb),
        [entry("A1", &["NOPE"]), entry("A2", &["NOPE"])]
    );
    assert_eq!(value(&wb, "A1"), number(42.0));
    assert_eq!(value(&wb, "A2"), number(42.0));
}

/// the reference forms: a name that is no call is no function, so it keeps
/// the error excel shows and a handler answers it.
#[test]
fn an_undefined_name_that_is_no_call_keeps_its_error() {
    let mut wb = book(&[
        ("A1", Some("NoSuchName*2"), number(42.0)),
        ("A2", Some("IFERROR(NoSuchName,0)"), number(42.0)),
        ("A3", Some("NoSuchName"), CellValue::Empty),
    ]);
    assert_eq!(listed(&mut wb), []);
    assert_eq!(value(&wb, "A1"), NAME);
    assert_eq!(value(&wb, "A2"), number(0.0));
    assert_eq!(value(&wb, "A3"), NAME);
}

#[test]
fn a_call_through_a_workbook_or_sheet_qualifier_keeps_excels_value() {
    for (formula, reported) in [
        ("[1]!Foo(C1)", "[1]!Foo"),
        ("IFERROR([1]!Foo(C1),0)", "[1]!Foo"),
        ("[12]!My.Fn(C1)+1", "[12]!My.Fn"),
        ("Sheet1!Foo(C1)", "Sheet1!Foo"),
        ("'[1]Add-in'!Foo(C1)", "'[1]Add-in'!Foo"),
        ("'Odd Sheet'!Foo(C1)", "'Odd Sheet'!Foo"),
    ] {
        let mut wb = book(&[
            ("A1", Some(formula), number(42.0)),
            ("B1", Some("A1*2"), number(5.0)),
            ("C1", None, number(1.0)),
        ]);
        assert_eq!(
            listed(&mut wb),
            [entry("A1", &[reported]), entry("B1", &[reported])],
            "{formula}"
        );
        assert_eq!(value(&wb, "A1"), number(42.0), "{formula}");
        assert_eq!(value(&wb, "B1"), number(5.0), "{formula}");
    }
}

/// `[1]Data!A1`, a reference, stays as it was: unparsed and kept.
#[test]
fn a_reference_into_another_workbook_is_no_call() {
    let mut wb = book(&[
        ("A1", Some("[1]Data!A1+1"), number(42.0)),
        ("A2", Some("[1]!Rate*2"), number(42.0)),
    ]);
    assert_eq!(listed(&mut wb), []);
    assert_eq!(value(&wb, "A1"), number(42.0));
    assert_eq!(value(&wb, "A2"), number(42.0));
}

#[test]
fn a_branch_never_taken_reaches_nothing() {
    for (formula, expected) in [
        ("IF(TRUE,1,NOPE(1))", 1.0),
        ("IFERROR(1,NOPE(1))", 1.0),
        ("CHOOSE(1,1,NOPE(1))", 1.0),
        ("SUM(IF({1,1},1,NOPE(1)))", 2.0),
    ] {
        let mut wb = book(&[("A1", Some(formula), number(42.0))]);
        assert_eq!(listed(&mut wb), [], "{formula}");
        assert_eq!(value(&wb, "A1"), number(expected), "{formula}");
    }
}

#[test]
fn a_kept_array_keeps_its_rectangle_and_its_readers() {
    let mut wb = book(&[
        ("A1", None, number(1.0)),
        ("A2", None, number(2.0)),
        ("B1", Some("NOPE(A1:A2)"), number(10.0)),
        ("B2", None, number(20.0)),
        ("C1", Some("B2+1"), number(5.0)),
    ]);
    wb.sheets[0].set_array_formula(
        CellRef::parse_a1("B1").unwrap(),
        CellRange::parse_a1("B1:B2").unwrap(),
    );
    assert_eq!(
        listed(&mut wb),
        [entry("B1", &["NOPE"]), entry("C1", &["NOPE"])]
    );
    assert_eq!(value(&wb, "B1"), number(10.0));
    assert_eq!(value(&wb, "B2"), number(20.0));
    assert_eq!(value(&wb, "C1"), number(5.0));
}

/// a formula reading through `INDIRECT` learns of the call only as it reads:
/// a kept value holds it, an answer does not.
#[test]
fn a_reader_through_a_computed_reference_tells_kept_from_answered() {
    let mut wb = book(&[
        ("A1", Some("NOPE(1)"), number(21.0)),
        ("B1", Some(r#"INDIRECT("A1")*2"#), number(5.0)),
        ("A2", Some("NOPE(1)"), CellValue::Empty),
        ("B2", Some(r#"INDIRECT("A2")"#), number(5.0)),
    ]);
    let reported = [
        entry("A1", &["NOPE"]),
        entry("B1", &["NOPE"]),
        entry("A2", &["NOPE"]),
        entry("B2", &["NOPE"]),
    ];
    assert_eq!(listed(&mut wb), reported);
    assert_eq!(value(&wb, "B1"), number(5.0));
    assert_eq!(value(&wb, "B2"), NAME);
}

/// an array formula with nothing saved answers over its rectangle, and a
/// formula reading a cell of that rectangle is listed.
#[test]
fn an_answered_array_lists_the_readers_of_its_rectangle() {
    let mut wb = book(&[
        ("B1", Some("IFERROR(NOPE(1),{1;2})"), CellValue::Empty),
        ("C1", Some("B2+1"), CellValue::Empty),
    ]);
    wb.sheets[0].set_array_formula(
        CellRef::parse_a1("B1").unwrap(),
        CellRange::parse_a1("B1:B2").unwrap(),
    );
    assert_eq!(
        listed(&mut wb),
        [entry("B1", &["NOPE"]), entry("C1", &["NOPE"])]
    );
    assert_eq!(value(&wb, "B1"), number(1.0));
    assert_eq!(value(&wb, "B2"), number(2.0));
    assert_eq!(value(&wb, "C1"), number(3.0));
}

/// the array's result outgrows the rectangle recorded for it, so only the
/// result says that C1 reads it.
fn outgrown_array() -> Workbook {
    let mut wb = book(&[
        ("B1", Some("IFERROR(NOPE(1),{1;2})"), CellValue::Empty),
        ("C1", Some("B2+1"), CellValue::Empty),
    ]);
    let anchor = CellRef::parse_a1("B1").unwrap();
    wb.sheets[0].set_array_formula(anchor, CellRange::new(anchor, anchor));
    wb
}

/// a gap never reshapes a rectangle, so `#NAME?` fills the recorded one, as
/// on main; C1 reads a cell the real result may cover and is listed for it.
#[test]
fn the_readers_of_an_answer_past_its_recorded_rectangle_are_listed() {
    let mut wb = outgrown_array();
    assert_eq!(
        listed(&mut wb),
        [entry("B1", &["NOPE"]), entry("C1", &["NOPE"])]
    );
    assert_eq!(value(&wb, "B1"), NAME);
    assert_eq!(value(&wb, "C1"), number(1.0));
}

/// the array itself only reads an answer, so it fills what its result does,
/// and a formula reading that is listed.
#[test]
fn an_array_reading_an_answer_lists_the_readers_of_its_result() {
    let mut wb = book(&[
        ("A1", Some("NOPE(1)"), CellValue::Empty),
        ("B1", Some("IFERROR(A1,{1;2})"), CellValue::Empty),
        ("C1", Some("B2+1"), CellValue::Empty),
    ]);
    let anchor = CellRef::parse_a1("B1").unwrap();
    wb.sheets[0].set_array_formula(anchor, CellRange::new(anchor, anchor));
    assert_eq!(
        listed(&mut wb),
        [
            entry("A1", &["NOPE"]),
            entry("B1", &["NOPE"]),
            entry("C1", &["NOPE"]),
        ]
    );
    assert_eq!(value(&wb, "B2"), number(2.0));
    assert_eq!(value(&wb, "C1"), number(3.0));
}

#[test]
fn opening_lists_the_readers_of_an_answer_past_its_recorded_rectangle() {
    let wb = outgrown_array();
    let graph = DepGraph::build(&wb);
    derive_unevaluated(&wb, &graph, None, None);
    let cells: Vec<_> = graph
        .unevaluated_cells()
        .iter()
        .map(|cell| (cell.cell.to_a1(), cell.functions.clone()))
        .collect();
    assert_eq!(cells, [entry("B1", &["NOPE"]), entry("C1", &["NOPE"])]);
}

/// A1 answers in the first round, before anything shows that `INDIRECT` reads
/// a cell an uncomputable call holds; the next round puts it back.
#[test]
fn an_answer_a_later_round_refuses_is_put_back() {
    let mut wb = book(&[
        (
            "A1",
            Some(r#"IFERROR(NOPE(1),INDIRECT("B1"))"#),
            CellValue::Empty,
        ),
        ("B1", Some(r#"WEBSERVICE("http://x")"#), CellValue::Empty),
    ]);
    assert_eq!(
        listed(&mut wb),
        [
            entry("A1", &["NOPE", "WEBSERVICE"]),
            entry("B1", &["WEBSERVICE"]),
        ]
    );
    assert_eq!(value(&wb, "A1"), CellValue::Empty);
}
