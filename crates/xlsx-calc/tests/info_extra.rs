//! ERROR.TYPE, INFO, ISFORMULA, ISNONTEXT, ISOMITTED, ISREF, SHEET, SHEETS and
//! TYPE against microsoft's function reference.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn boolean(value: bool) -> CellValue {
    CellValue::Bool { value }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn book() -> Workbook {
    Workbook {
        sheets: vec![
            Sheet::new("Sheet1"),
            Sheet::new("My Data"),
            Sheet::new("Stuff"),
        ],
        ..Workbook::default()
    }
}

fn put(book: &mut Workbook, address: &str, cell: Cell) {
    book.sheets[0].set_cell(CellRef::parse_a1(address).unwrap(), cell);
}

fn put_value(book: &mut Workbook, address: &str, value: CellValue) {
    put(
        book,
        address,
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn put_formula(book: &mut Workbook, address: &str, formula: &str) {
    put(
        book,
        address,
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
}

fn put_cached(book: &mut Workbook, address: &str, formula: &str, cached: CellValue) {
    put(
        book,
        address,
        Cell {
            formula: Some(formula.into()),
            value: cached,
            ..Cell::default()
        },
    );
}

fn value(book: &Workbook, address: &str) -> CellValue {
    book.sheets[0]
        .cell(CellRef::parse_a1(address).unwrap())
        .unwrap()
        .value
        .clone()
}

/// evaluate each formula against `book`, one result cell per formula.
fn check(mut book: Workbook, cases: &[(&str, CellValue)]) {
    for (i, (formula, _)) in cases.iter().enumerate() {
        put_formula(&mut book, &format!("Z{}", i + 1), formula);
    }
    rebuild_and_recalc_all(&mut book, None);
    for (i, (formula, want)) in cases.iter().enumerate() {
        assert_eq!(&value(&book, &format!("Z{}", i + 1)), want, "{formula}");
    }
}

#[test]
fn error_type_numbers_each_error() {
    let mut book = book();
    put_value(&mut book, "A2", error(ErrorValue::Null));
    put_value(&mut book, "A3", error(ErrorValue::Div0));
    let doc = "IF(ERROR.TYPE(A2)<3,CHOOSE(ERROR.TYPE(A2),\"Ranges do not intersect\",\"The divisor is zero\"))";
    let doc_div = "IF(ERROR.TYPE(A3)<3,CHOOSE(ERROR.TYPE(A3),\"Ranges do not intersect\",\"The divisor is zero\"))";
    check(
        book,
        &[
            (doc, text("Ranges do not intersect")),
            (doc_div, text("The divisor is zero")),
            ("ERROR.TYPE(#NULL!)", number(1.0)),
            ("ERROR.TYPE(1/0)", number(2.0)),
            ("ERROR.TYPE(#VALUE!)", number(3.0)),
            ("ERROR.TYPE(#REF!)", number(4.0)),
            ("ERROR.TYPE(#NAME?)", number(5.0)),
            ("ERROR.TYPE(SQRT(-1))", number(6.0)),
            ("ERROR.TYPE(NA())", number(7.0)),
            ("_xlfn.ERROR.TYPE(#N/A)", number(7.0)),
            ("ERROR.TYPE(1)", error(ErrorValue::NA)),
            ("ERROR.TYPE(\"#N/A\")", error(ErrorValue::NA)),
            ("ERROR.TYPE(C40)", error(ErrorValue::NA)),
            ("ERROR.TYPE()", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn info_answers_the_offline_types() {
    check(
        book(),
        &[
            ("INFO(\"recalc\")", text("Automatic")),
            ("INFO(\"RECALC\")", text("Automatic")),
            ("INFO(\"memavail\")", error(ErrorValue::NA)),
            ("INFO(\"memused\")", error(ErrorValue::NA)),
            ("INFO(\"totmem\")", error(ErrorValue::NA)),
            ("INFO(\"numfile\")", number(3.0)),
            ("INFO(\"origin\")", text("$A:$A$1")),
            ("INFO(\"system\")", error(ErrorValue::NA)),
            ("INFO(\"bogus\")", error(ErrorValue::Value)),
            ("INFO(1)", error(ErrorValue::Value)),
            ("INFO(1/0)", error(ErrorValue::Div0)),
            ("INFO()", error(ErrorValue::Value)),
        ],
    );
}

/// the machine and window facts are engine gaps: a cell keeps what excel
/// cached for them.
#[test]
fn info_machine_facts_keep_the_value_excel_cached() {
    let mut book = book();
    put_cached(&mut book, "B1", "INFO(\"system\")", text("pcdos"));
    put_cached(
        &mut book,
        "B2",
        "INFO(\"osversion\")",
        text("Windows (64-bit) NT 10.00"),
    );
    put_cached(
        &mut book,
        "B3",
        "INFO(\"directory\")",
        text("C:\\Users\\me\\Documents\\"),
    );
    put_cached(&mut book, "B4", "INFO(\"release\")", text("16.0"));
    put_cached(&mut book, "B5", "INFO(\"numfile\")", number(7.0));
    put_cached(&mut book, "B6", "INFO(\"origin\")", text("$A:$C$40"));
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "B1"), text("pcdos"));
    assert_eq!(value(&book, "B2"), text("Windows (64-bit) NT 10.00"));
    assert_eq!(value(&book, "B3"), text("C:\\Users\\me\\Documents\\"));
    assert_eq!(value(&book, "B4"), text("16.0"));
    assert_eq!(value(&book, "B5"), number(7.0));
    assert_eq!(value(&book, "B6"), text("$A:$C$40"));
}

#[test]
fn isformula_reads_the_referenced_cell() {
    let mut book = book();
    put_formula(&mut book, "A1", "1+1");
    put_value(&mut book, "A2", number(2.0));
    put_formula(&mut book, "A3", "TODAY()");
    check(
        book,
        &[
            ("ISFORMULA(A3)", boolean(true)),
            ("ISFORMULA(A1)", boolean(true)),
            ("_xlfn.ISFORMULA(A2)", boolean(false)),
            ("ISFORMULA(A1:A2)", boolean(true)),
            ("ISFORMULA(A2:A3)", boolean(false)),
            ("ISFORMULA(C40)", boolean(false)),
            ("ISFORMULA(INDIRECT(\"A1\"))", boolean(true)),
            ("ISFORMULA(1)", error(ErrorValue::Value)),
            ("ISFORMULA(\"A1\")", error(ErrorValue::Value)),
            ("ISFORMULA(1/0)", error(ErrorValue::Div0)),
            ("ISFORMULA()", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn isnontext_is_true_for_everything_but_text() {
    let mut book = book();
    put_value(&mut book, "A1", text("Region1"));
    put_value(&mut book, "A2", number(330.92));
    check(
        book,
        &[
            ("ISNONTEXT(A1)", boolean(false)),
            ("ISNONTEXT(A2)", boolean(true)),
            ("ISNONTEXT(C40)", boolean(true)),
            ("ISNONTEXT(19)", boolean(true)),
            ("ISNONTEXT(TRUE)", boolean(true)),
            ("ISNONTEXT(1/0)", boolean(true)),
            ("ISNONTEXT(\"\")", boolean(false)),
            ("ISNONTEXT(\"19\")", boolean(false)),
            ("SUMPRODUCT(--ISNONTEXT({1,\"a\",TRUE}))", number(2.0)),
            ("ISNONTEXT()", error(ErrorValue::Value)),
        ],
    );
}

/// the engine binds every lambda parameter, so inside a callback nothing is
/// omitted.
#[test]
fn isomitted_sees_bound_parameters() {
    check(
        book(),
        &[
            (
                "_xlfn.REDUCE(0,{1,2,3},_xlfn.LAMBDA(a,b,a+ISOMITTED(b)))",
                number(0.0),
            ),
            ("ISOMITTED(1)", boolean(false)),
            ("_xlfn.ISOMITTED(A1)", boolean(false)),
            ("ISOMITTED()", error(ErrorValue::Value)),
            ("ISOMITTED(1,2)", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn isref_tells_references_from_values() {
    check(
        book(),
        &[
            ("ISREF(G8)", boolean(true)),
            ("ISREF(A1:B2)", boolean(true)),
            ("ISREF('My Data'!C3)", boolean(true)),
            ("ISREF((A1,B2))", boolean(true)),
            ("ISREF(A:A)", boolean(true)),
            ("ISREF(OFFSET(A1,1,1))", boolean(true)),
            ("ISREF(INDIRECT(\"A1\"))", boolean(true)),
            ("ISREF(INDIRECT(\"not a ref\"))", boolean(false)),
            ("ISREF(1)", boolean(false)),
            ("ISREF(\"A1\")", boolean(false)),
            ("ISREF(1/0)", boolean(false)),
            ("ISREF({1,2})", boolean(false)),
            ("ISREF()", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn sheet_numbers_and_sheets_counts() {
    let mut book = book();
    let at = CellRef::parse_a1("B2").unwrap();
    book.sheets[1].set_cell(
        at,
        Cell {
            formula: Some("SHEET()".into()),
            ..Cell::default()
        },
    );
    check(
        book.clone(),
        &[
            ("SHEET()", number(1.0)),
            ("_xlfn.SHEET(A1)", number(1.0)),
            ("SHEET('My Data'!A1:B4)", number(2.0)),
            ("SHEET(\"Stuff\")", number(3.0)),
            ("SHEET(\"stuff\")", number(3.0)),
            ("SHEET(\"Nope\")", error(ErrorValue::NA)),
            ("SHEET(INDIRECT(\"'My Data'!A1\"))", number(2.0)),
            ("SHEET(1/0)", error(ErrorValue::Div0)),
            ("SHEET(1,2)", error(ErrorValue::Value)),
            ("SHEETS()", number(3.0)),
            ("_xlfn.SHEETS(A1)", number(1.0)),
            ("SHEETS('Stuff'!A1:C9)", number(1.0)),
            ("SHEETS((A1,B2))", number(1.0)),
            ("SHEETS(1/0)", error(ErrorValue::Div0)),
            ("SHEETS(1)", error(ErrorValue::Value)),
            ("SHEETS(A1,A2)", error(ErrorValue::Value)),
        ],
    );
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(book.sheets[1].cell(at).unwrap().value, number(2.0));
}

#[test]
fn type_classifies_values() {
    let mut book = book();
    put_value(&mut book, "A2", text("Smith"));
    put_value(&mut book, "A3", boolean(true));
    put_value(&mut book, "A4", error(ErrorValue::Ref));
    check(
        book,
        &[
            ("TYPE(A2)", number(2.0)),
            ("TYPE(\"MR. \"&A2)", number(2.0)),
            ("TYPE(2+A2)", number(16.0)),
            ("TYPE({1,2;3,4})", number(64.0)),
            ("TYPE(1)", number(1.0)),
            ("TYPE(C40)", number(1.0)),
            ("TYPE(A3)", number(4.0)),
            ("TYPE(A4)", number(16.0)),
            ("TYPE(1/0)", number(16.0)),
            ("TYPE(A2:A3)", number(64.0)),
            ("TYPE(_xlfn.SEQUENCE(3))", number(64.0)),
            ("TYPE(_xlfn.LAMBDA(x,x))", number(128.0)),
            ("TYPE()", error(ErrorValue::Value)),
        ],
    );
}
