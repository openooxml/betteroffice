//! excel's union operator: a `,` inside parentheses joins references into one
//! multi-area reference, read by the aggregates that accept several areas.

use xlsx_calc::{EvalContext, evaluate, parse_formula, rebuild_and_recalc_all, references};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn put(sheet: &mut Sheet, a1: &str, value: CellValue, formula: Option<&str>) {
    sheet.set_cell(
        CellRef::parse_a1(a1).unwrap(),
        Cell {
            value,
            formula: formula.map(str::to_string),
            style: None,
        },
    );
}

/// Sheet1: A1:A5 = 10..50, B1:B2 text, C1:C5 = 1..5, E1:E4 = 1..4.
/// Other: A1:A2 = 7, 8.
fn fixture() -> Workbook {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    for i in 1..=5 {
        put(&mut s, &format!("A{i}"), n(10.0 * i as f64), None);
        put(&mut s, &format!("C{i}"), n(i as f64), None);
    }
    for i in 1..=4 {
        put(&mut s, &format!("E{i}"), n(i as f64), None);
    }
    for (a1, text) in [("B1", "x"), ("B2", "y")] {
        put(&mut s, a1, CellValue::Text { value: text.into() }, None);
    }
    wb.sheets.push(s);
    let mut other = Sheet::new("Other");
    put(&mut other, "A1", n(7.0), None);
    put(&mut other, "A2", n(8.0), None);
    wb.sheets.push(other);
    wb
}

#[track_caller]
fn check(cases: &[(&str, CellValue)]) {
    let wb = fixture();
    let ctx = EvalContext::new(&wb, SheetId(0));
    for (src, want) in cases {
        let expr = parse_formula(src).unwrap_or_else(|e| panic!("{src}: {e:?}"));
        assert_eq!(&evaluate(&expr, &ctx), want, "{src}");
    }
}

#[test]
fn union_formulas_print_back_as_written() {
    for src in [
        "RANK(Q17,(C17:Q17,U17:AG17),1)",
        "SUM((A1,B1:B2),C1)",
        "SUM(((A1,B1),C1))",
        "SUM((Other!A1:A2,'My Sheet'!B1))",
        "(A1,B1)+1",
    ] {
        let expr = parse_formula(src).unwrap_or_else(|e| panic!("{src}: {e:?}"));
        assert_eq!(expr.to_formula(), src);
    }
}

#[test]
fn a_union_depends_on_every_area() {
    let expr = parse_formula("RANK(Q17,(C17:Q17,U17:AG17),1)").unwrap();
    let read: Vec<String> = references(&expr)
        .into_iter()
        .map(|(_, range)| range.to_a1())
        .collect();
    assert_eq!(read, ["Q17", "C17:Q17", "U17:AG17"]);
}

/// the union reads 1,2,3,4,5 and 1,2,3,4, so every value below 5 is tied.
#[test]
fn rank_reads_every_area_of_a_union() {
    check(&[
        ("RANK(3,(C1:C5,E1:E4))", n(4.0)),
        ("RANK(3,(C1:C5,E1:E4),1)", n(5.0)),
        ("RANK(1,(C1:C5,E1:E4))", n(8.0)),
        ("RANK(5,(C1:C5,E1:E4),1)", n(9.0)),
        ("_xlfn.RANK.EQ(3,(C1:C5,E1:E4),1)", n(5.0)),
        ("_xlfn.RANK.AVG(3,(C1:C5,E1:E4))", n(4.5)),
        ("_xlfn.RANK.AVG(3,(C1:C5,E1:E4),1)", n(5.5)),
        ("_xlfn.RANK.AVG(5,C1:C5)", n(1.0)),
        ("RANK(9,(C1:C5,E1:E4))", e(ErrorValue::NA)),
    ]);
}

#[test]
fn aggregates_read_every_area_of_a_union() {
    check(&[
        ("SUM((A1:A2,C1:C3))", n(36.0)),
        ("SUM((A1,C1),E1)", n(12.0)),
        ("SUM(((A1,C1),E1))", n(12.0)),
        ("SUM((Other!A1:A2,Other!A1))", n(22.0)),
        ("COUNT((A1:A5,B1:B2))", n(5.0)),
        ("COUNTA((A1:A2,B1:B2))", n(4.0)),
        ("AVERAGE((A1,C1))", n(5.5)),
        ("MIN((A1:A5,C2))", n(2.0)),
        ("MAX((C1:C5,E1:E4))", n(5.0)),
        ("LARGE((A1:A2,C1:C5),2)", n(10.0)),
        ("SMALL((A1:A2,C4:C5),3)", n(10.0)),
        ("MEDIAN((C1:C3,E4))", n(2.5)),
        ("AREAS((A1,B1:B2,C1))", n(3.0)),
        ("AREAS(((A1,B1),C1))", n(3.0)),
        ("AREAS(A1:B2)", n(1.0)),
    ]);
}

#[test]
fn a_union_used_as_a_value_is_a_value_error() {
    check(&[
        ("(A1,C1)", e(ErrorValue::Value)),
        ("(A1,C1)+1", e(ErrorValue::Value)),
        ("ABS((A1,C1))", e(ErrorValue::Value)),
        ("SUM((A1,Other!A1))", e(ErrorValue::Value)),
        ("SUM((A1,nosuchname))", e(ErrorValue::Name)),
    ]);
}

#[test]
fn a_full_recalc_orders_a_union_after_every_area() {
    let mut wb = Workbook::default();
    let mut s = Sheet::new("Sheet1");
    put(&mut s, "A1", CellValue::Empty, Some("SUM((H1,J1:J2))"));
    put(
        &mut s,
        "A2",
        CellValue::Empty,
        Some("RANK(H1,(H1,J1:J2),1)"),
    );
    put(&mut s, "H1", CellValue::Empty, Some("1+1"));
    put(&mut s, "J1", CellValue::Empty, Some("2+2"));
    put(&mut s, "J2", CellValue::Empty, Some("H1*10"));
    wb.sheets.push(s);
    rebuild_and_recalc_all(&mut wb, None);
    let value = |a1: &str| {
        wb.sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::parse_a1(a1).unwrap())
            .unwrap()
            .value
            .clone()
    };
    assert_eq!(value("A1"), n(26.0));
    assert_eq!(value("A2"), n(1.0));
}
