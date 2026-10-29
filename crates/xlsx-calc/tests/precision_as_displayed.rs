use xlsx_calc::{rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellRef, CellValue, Sheet, SheetId, Workbook, Xf};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn put(book: &mut Workbook, address: &str, formula: &str, format: &str) {
    let id = 164 + book.styles.num_fmts.len() as u16;
    let style = book.styles.cell_xfs.len() as u32;
    book.styles.num_fmts.push((id, format.into()));
    book.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(id),
        ..Xf::default()
    });
    book.sheets[0].set_cell(
        CellRef::parse_a1(address).unwrap(),
        Cell {
            formula: Some(formula.into()),
            style: Some(style),
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

fn workbook() -> Workbook {
    Workbook {
        sheets: vec![Sheet::new("Sheet1")],
        ..Workbook::default()
    }
}

#[test]
fn displayed_precision_reaches_dependents_and_incremental_recalculation() {
    let mut book = workbook();
    book.full_precision = false;
    put(&mut book, "A1", "6.102019382885104", "0.0");
    put(&mut book, "B1", "A1*10", "General");
    let (mut graph, _) = rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "A1"), number(6.1));
    assert_eq!(value(&book, "B1"), number(61.0));
    book.sheets[0].cell_mut(CellRef::new(0, 0)).unwrap().formula = Some("6.26".into());
    graph.set_formula(SheetId(0), CellRef::new(0, 0), Some("6.26"));
    recalc_after(
        &mut book,
        &mut graph,
        &[(SheetId(0), CellRef::new(0, 0))],
        None,
    );
    assert_eq!(value(&book, "A1"), number(6.3));
    assert_eq!(value(&book, "B1"), number(63.0));
}

#[test]
fn full_precision_is_the_default() {
    let mut book = workbook();
    assert!(book.full_precision);
    put(&mut book, "A1", "6.102019382885104", "0.0");
    put(&mut book, "B1", "A1*10", "General");
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "A1"), number(6.102019382885104));
    assert_eq!(value(&book, "B1"), number(6.102019382885104 * 10.0));
}

#[test]
fn displayed_precision_uses_numeric_format_sections() {
    for (formula, format, expected) in [
        ("6.9", "0", 7.0),
        ("0.126", "0%", 0.13),
        ("0.1265", "0.0%", 0.127),
        ("12550", "#,##0,", 13000.0),
        ("12550", "#,##0.0,", 12600.0),
        ("12550", "#,##0,.0", 12600.0),
        ("-6.125", "0.0;(0.00)", -6.13),
        ("6.125", "0.0;(0.00)", 6.1),
        ("1.25", "0.0", 1.3),
        ("-1.25", "0.0", -1.3),
        ("1234.567", "#,##0.00", 1234.57),
        ("6.125", "0.00_);(0.00)", 6.13),
        ("6.125", "\"USD\" 0.00", 6.13),
        ("6.125", "[$$-409]0.00", 6.13),
        ("12345.6", "0.00E+00", 12300.0),
        ("-0.0123456", "0.00E+00", -0.0123),
        ("6.125", "[Red]0.0", 6.1),
    ] {
        let mut book = workbook();
        book.full_precision = false;
        put(&mut book, "A1", formula, format);
        rebuild_and_recalc_all(&mut book, None);
        assert_eq!(value(&book, "A1"), number(expected), "{formula}: {format}");
    }
}

#[test]
fn displayed_precision_leaves_unclassified_formats_and_constants_unchanged() {
    for format in [
        "General",
        "m/d/yy",
        "hh:mm:ss.00",
        "[h]:mm",
        "[<1000]0.0;0.00",
        "##0.0E+00",
        "\"hidden\"",
        "0.0suffix",
        "0.0\\",
        "0.0_",
        "0.0*",
        "@",
        "# ?/?",
        "0.0.0",
        "[unknown]0.0",
        "0.0\"unfinished",
        "[$USD0.0",
    ] {
        let mut book = workbook();
        book.full_precision = false;
        put(&mut book, "A1", "6.125", format);
        put(&mut book, "B1", "6.125", "0.0");
        let constant = book.sheets[0].cell_mut(CellRef::new(0, 1)).unwrap();
        constant.formula = None;
        constant.value = number(6.125);
        rebuild_and_recalc_all(&mut book, None);
        assert_eq!(value(&book, "A1"), number(6.125), "{format}");
        assert_eq!(value(&book, "B1"), number(6.125));
    }
}

#[test]
fn displayed_precision_rounds_spills_by_each_cells_format() {
    let mut book = workbook();
    book.full_precision = false;
    put(&mut book, "A1", "{6.125,6.125}", "0.0");
    book.sheets[0].set_array_formula(
        CellRef::new(0, 0),
        xlsx_model::CellRange::parse_a1("A1:B1").unwrap(),
    );
    put(&mut book, "B1", "0", "0.00");
    book.sheets[0].cell_mut(CellRef::new(0, 1)).unwrap().formula = None;
    put(&mut book, "C1", "B1*100", "General");
    let (mut graph, _) = rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "A1"), number(6.1));
    assert_eq!(value(&book, "B1"), number(6.13));
    assert_eq!(value(&book, "C1"), number(613.0));
    book.sheets[0].cell_mut(CellRef::new(0, 0)).unwrap().formula = Some("{7.255,7.255}".into());
    graph.set_formula(SheetId(0), CellRef::new(0, 0), Some("{7.255,7.255}"));
    recalc_after(
        &mut book,
        &mut graph,
        &[(SheetId(0), CellRef::new(0, 0))],
        None,
    );
    assert_eq!(value(&book, "A1"), number(7.3));
    assert_eq!(value(&book, "B1"), number(7.26));
    assert_eq!(value(&book, "C1"), number(726.0));
}

#[test]
fn displayed_precision_preserves_non_numeric_results() {
    let mut book = workbook();
    book.full_precision = false;
    for (address, formula) in [("A1", "TRUE()"), ("B1", "\"text\""), ("C1", "1/0")] {
        put(&mut book, address, formula, "0.0");
    }
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "A1"), CellValue::Bool { value: true });
    assert_eq!(
        value(&book, "B1"),
        CellValue::Text {
            value: "text".into()
        }
    );
    assert_eq!(
        value(&book, "C1"),
        CellValue::Error {
            value: xlsx_model::ErrorValue::Div0
        }
    );
}

#[test]
fn displayed_precision_normalizes_negative_zero() {
    for cached in [-1.0, -0.0] {
        let mut book = workbook();
        book.full_precision = false;
        put(&mut book, "A1", "-0.04", "0.0");
        book.sheets[0].cell_mut(CellRef::new(0, 0)).unwrap().value = number(cached);
        rebuild_and_recalc_all(&mut book, None);
        let CellValue::Number { value } = value(&book, "A1") else {
            panic!("expected number");
        };
        assert_eq!(value.to_bits(), 0.0_f64.to_bits());
    }
}

#[test]
fn displayed_precision_does_not_touch_unrecalculated_caches() {
    let mut book = workbook();
    book.full_precision = false;
    put(&mut book, "A1", "6.125", "0.0");
    put(&mut book, "B1", "6.125", "0.0");
    let (mut graph, _) = rebuild_and_recalc_all(&mut book, None);
    book.sheets[0].cell_mut(CellRef::new(0, 1)).unwrap().value = number(6.125);
    recalc_after(
        &mut book,
        &mut graph,
        &[(SheetId(0), CellRef::new(0, 0))],
        None,
    );
    assert_eq!(value(&book, "A1"), number(6.1));
    assert_eq!(value(&book, "B1"), number(6.125));
}

#[test]
fn displayed_precision_rounds_values_settled_in_a_range_cycle() {
    let mut book = workbook();
    book.full_precision = false;
    put(&mut book, "A1", "SUM(B1:B2)", "0.0");
    put(&mut book, "B1", "IF(FALSE,A1,6.125)", "0.00");
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "B1"), number(6.13));
    assert_eq!(value(&book, "A1"), number(6.1));
}
