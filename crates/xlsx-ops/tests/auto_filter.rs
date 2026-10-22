use xlsx_model::{CellRange, Sheet, SheetId, Workbook};
use xlsx_ops::{Op, Provenance, Transaction, UndoStack, apply};

#[test]
fn null_without_a_filter_is_a_no_op_with_no_inverse() {
    let mut workbook = Workbook::default();
    workbook.sheets.push(Sheet::new("Data 7"));
    let before = workbook.clone();
    let inverse = apply(&mut workbook, &Op::SetAutoFilter { sheet: SheetId(0), range: None }).unwrap();
    assert!(inverse.0.is_empty());
    assert_eq!(workbook, before);
    let mut stack = UndoStack::new();
    stack.commit(&mut workbook, &Transaction::new(vec![Op::SetAutoFilter {
        sheet: SheetId(0), range: None,
    }], Provenance::User)).unwrap();
    assert_eq!(stack.undo_depth(), 0);
}

#[test]
fn create_remove_and_inverse_restore_names_and_filter() {
    let mut workbook = Workbook::default();
    workbook.sheets.push(Sheet::new("Data 7"));
    let before = workbook.clone();
    let inverse = apply(&mut workbook, &Op::SetAutoFilter {
        sheet: SheetId(0), range: Some(CellRange::parse_a1("A3:F30").unwrap()),
    }).unwrap();
    assert_eq!(workbook.sheets[0].auto_filter.as_ref().unwrap().range.to_a1(), "A3:F30");
    assert_eq!(workbook.defined_names[0].formula, "'Data 7'!$A$3:$F$30");
    assert!(workbook.defined_names[0].hidden);
    let filtered = workbook.clone();
    let removed = apply(&mut workbook, &Op::SetAutoFilter { sheet: SheetId(0), range: None }).unwrap();
    assert_eq!(workbook, before);
    for op in removed.0 { apply(&mut workbook, &op).unwrap(); }
    assert_eq!(workbook, filtered);
    for op in inverse.0 { apply(&mut workbook, &op).unwrap(); }
    assert_eq!(workbook, before);
}

#[test]
fn unsafe_ranges_and_duplicate_sources_refuse_without_changes() {
    for (range, duplicate, table) in [
        ("B3:F30", false, false),
        ("A3:B30", false, false),
        ("A3:F30", true, false),
        ("A3:F30", false, true),
    ] {
        let mut workbook = Workbook::default();
        let mut sheet = Sheet::new("Data 7");
        sheet.auto_filter = Some(xlsx_model::AutoFilter {
            range: CellRange::parse_a1("A3:F30").unwrap(),
            raw_xml: Some("<autoFilter ref=\"A3:F30\"><filterColumn colId=\"5\"/></autoFilter>".into()),
            source_count: if duplicate { 2 } else { 1 },
            filter_columns: vec![5],
            has_sort_state: false,
        });
        workbook.sheets.push(sheet);
        if table {
            workbook.tables.push(xlsx_model::Table {
                name: "Items".into(), sheet: SheetId(0),
                range: CellRange::parse_a1("E3:F30").unwrap(),
                header_rows: 1, totals_rows: 0, columns: vec!["A".into(), "B".into()],
            });
        }
        let before = workbook.clone();
        assert!(apply(&mut workbook, &Op::SetAutoFilter {
            sheet: SheetId(0), range: Some(CellRange::parse_a1(range).unwrap()),
        }).is_err());
        assert_eq!(workbook, before);
        if duplicate {
            assert!(apply(&mut workbook, &Op::SetAutoFilter { sheet: SheetId(0), range: None }).is_err());
            assert_eq!(workbook, before);
        }
    }
}

#[test]
fn wire_range_is_a_string_or_null() {
    let op: Op = serde_json::from_str(r#"{"type":"setAutoFilter","sheet":0,"range":"A3:F30"}"#).unwrap();
    assert!(matches!(op, Op::SetAutoFilter { range: Some(_), .. }));
    let op: Op = serde_json::from_str(r#"{"type":"setAutoFilter","sheet":0,"range":null}"#).unwrap();
    assert!(matches!(op, Op::SetAutoFilter { range: None, .. }));
    assert_eq!(serde_json::to_value(op).unwrap()["range"], serde_json::Value::Null);
}
