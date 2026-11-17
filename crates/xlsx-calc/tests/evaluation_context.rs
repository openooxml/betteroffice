use xlsx_calc::{EvalContext, evaluate, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, Sheet, SheetId, Workbook};

#[test]
fn a_restarted_context_evaluates_with_a_fresh_budget() {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    for row in 0..3 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    workbook.sheets.push(sheet);
    let context = EvalContext::new(&workbook, SheetId(0));
    let full = context.remaining_cells();
    let sum = parse_formula("SUM(A1:A3)").unwrap();
    assert_eq!(evaluate(&sum, &context), CellValue::Number { value: 3.0 });
    assert!(context.remaining_cells() < full);
    assert!(!context.exhausted());
    evaluate(&parse_formula("SUM(A1:ZZ100000)").unwrap(), &context);
    assert!(context.exhausted());
    context.restart();
    assert!(!context.exhausted());
    assert_eq!(context.remaining_cells(), full);
    assert_eq!(evaluate(&sum, &context), CellValue::Number { value: 3.0 });
}
