//! End-to-end demo: full recalc at open, then an incremental recalc after an edit.

use xlsx_calc::{rebuild_and_recalc_all, recalc_after};
use xlsx_model::{Cell, CellRef, CellValue, Sheet, SheetId, Workbook};

fn number(v: f64) -> Cell {
    Cell {
        value: CellValue::Number { value: v },
        formula: None,
        style: None,
    }
}

fn formula(src: &str) -> Cell {
    Cell {
        value: CellValue::Empty,
        formula: Some(src.into()),
        style: None,
    }
}

fn render(wb: &Workbook, sheet: SheetId, a1: &str) -> String {
    let at = CellRef::parse_a1(a1).unwrap();
    let cell = wb.sheet(sheet).unwrap().cell(at);
    let shown = match cell.map(|c| &c.value) {
        Some(CellValue::Number { value }) => value.to_string(),
        Some(CellValue::Text { value }) => value.clone(),
        Some(CellValue::Bool { value }) => value.to_string().to_uppercase(),
        Some(other) => format!("{other:?}"),
        None => "(empty)".into(),
    };
    match cell.and_then(|c| c.formula.as_deref()) {
        Some(f) => format!("{a1} = {f} -> {shown}"),
        None => format!("{a1} = {shown}"),
    }
}

fn main() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let s = SheetId(0);
    let sheet = wb.sheet_mut(s).unwrap();

    sheet.set_cell(CellRef::parse_a1("A1").unwrap(), number(3.0));
    sheet.set_cell(CellRef::parse_a1("A2").unwrap(), number(4.0));
    sheet.set_cell(CellRef::parse_a1("A3").unwrap(), formula("A1*A2"));
    sheet.set_cell(CellRef::parse_a1("B1").unwrap(), formula("SUM(A1:A3)"));
    sheet.set_cell(
        CellRef::parse_a1("B2").unwrap(),
        formula("IF(B1>=19,\"big\",\"small\")"),
    );

    let (mut graph, opened) = rebuild_and_recalc_all(&mut wb, None);
    println!("open: {} formulas evaluated", opened.changed.len());
    for a1 in ["A1", "A2", "A3", "B1", "B2"] {
        println!("  {}", render(&wb, s, a1));
    }

    let a1 = CellRef::parse_a1("A1").unwrap();
    wb.sheet_mut(s).unwrap().set_cell(a1, number(10.0));
    let edited = recalc_after(&mut wb, &mut graph, &[(s, a1)], None);
    println!("edit A1=10: {} cells changed", edited.changed.len());
    for a1 in ["A1", "A2", "A3", "B1", "B2"] {
        println!("  {}", render(&wb, s, a1));
    }
}
