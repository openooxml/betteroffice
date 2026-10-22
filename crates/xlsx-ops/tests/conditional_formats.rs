use xlsx_model::{CellRange, DifferentialFormat, Sheet, SheetId, Workbook};
use xlsx_ops::{Op, Provenance, Transaction, UndoStack, apply};

fn wb() -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    wb
}

fn add(range: &str, kind: &str, operator: Option<&str>, formulas: &[&str]) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1(range).unwrap(),
        kind: kind.into(),
        operator: operator.map(str::to_owned),
        formulas: formulas.iter().map(|f| f.to_string()).collect(),
        format: DifferentialFormat {
            font_color: Some("#FF0000".into()),
            bold: Some(true),
            fill: Some("#00FF00".into()),
        },
    }
}

#[test]
fn add_cell_is_and_expression_assigns_priority_above_existing() {
    let mut wb = wb();
    apply(
        &mut wb,
        &add("A1:A5", "cellIs", Some("greaterThan"), &["5"]),
    )
    .unwrap();
    apply(&mut wb, &add("B1:B5", "expression", None, &["=A1>5"])).unwrap();
    let groups = &wb.sheets[0].conditional_formats;
    assert_eq!(groups.len(), 2);
    assert_eq!(groups[0].sqref, "A1:A5");
    assert_eq!(groups[0].rules[0].priority, 1);
    assert_eq!(groups[0].rules[0].formulas, ["5"]);
    assert_eq!(groups[1].rules[0].priority, 2);
    assert_eq!(groups[1].rules[0].formulas, ["A1>5"]);
    assert_ne!(groups[0].rules[0].id, groups[1].rules[0].id);
    assert_eq!(groups[0].rules[0].format.bold, Some(true));
}

#[test]
fn multiple_rules_share_one_range_group() {
    let mut wb = wb();
    apply(
        &mut wb,
        &add("A1:A5", "cellIs", Some("greaterThan"), &["5"]),
    )
    .unwrap();
    apply(&mut wb, &add("A1:A5", "cellIs", Some("lessThan"), &["1"])).unwrap();
    assert_eq!(wb.sheets[0].conditional_formats.len(), 1);
    assert_eq!(wb.sheets[0].conditional_formats[0].rules.len(), 2);
}

#[test]
fn arity_and_operator_are_validated() {
    let mut wb = wb();
    assert!(apply(&mut wb, &add("A1", "expression", None, &["A1", "B1"])).is_err());
    assert!(apply(&mut wb, &add("A1", "cellIs", Some("between"), &["1"])).is_err());
    assert!(apply(&mut wb, &add("A1", "cellIs", Some("between"), &["1", "2"])).is_ok());
    assert!(apply(&mut wb, &add("A1", "cellIs", Some("bogus"), &["1"])).is_err());
    assert!(apply(&mut wb, &add("A1", "cellIs", None, &["1"])).is_err());
    assert!(apply(&mut wb, &add("A1", "colorScale", None, &["1"])).is_err());
    assert!(apply(&mut wb, &add("A1", "cellIs", Some("equal"), &[""])).is_err());
    assert!(
        wb.sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count()
            == 1
    );
}

#[test]
fn delete_by_id_and_undo_redo_restore() {
    let mut wb = wb();
    let mut stack = UndoStack::new();
    stack
        .commit(
            &mut wb,
            &Transaction::new(
                vec![add("A1:A5", "cellIs", Some("greaterThan"), &["5"])],
                Provenance::User,
            ),
        )
        .unwrap();
    let id = wb.sheets[0].conditional_formats[0].rules[0].id.clone();
    stack
        .commit(
            &mut wb,
            &Transaction::new(
                vec![Op::RemoveConditionalFormat {
                    sheet: SheetId(0),
                    id: id.clone(),
                }],
                Provenance::User,
            ),
        )
        .unwrap();
    assert!(wb.sheets[0].conditional_formats.is_empty());
    assert!(
        apply(
            &mut wb,
            &Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id
            }
        )
        .is_err()
    );
    stack.undo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].conditional_formats[0].rules.len(), 1);
    stack.undo(&mut wb).unwrap();
    assert!(wb.sheets[0].conditional_formats.is_empty());
    stack.redo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].conditional_formats[0].rules.len(), 1);
}

#[test]
fn bold_false_and_empty_format_facets_are_kept_as_given() {
    let mut wb = wb();
    let op = Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1").unwrap(),
        kind: "cellIs".into(),
        operator: Some("equal".into()),
        formulas: vec!["1".into()],
        format: DifferentialFormat {
            font_color: None,
            bold: Some(false),
            fill: None,
        },
    };
    apply(&mut wb, &op).unwrap();
    let format = &wb.sheets[0].conditional_formats[0].rules[0].format;
    assert_eq!(format.bold, Some(false));
    assert_eq!(format.font_color, None);
}

#[test]
fn bad_color_is_refused() {
    let mut wb = wb();
    let mut op = add("A1", "cellIs", Some("equal"), &["1"]);
    if let Op::AddConditionalFormat { format, .. } = &mut op {
        format.font_color = Some("red".into());
    }
    assert!(apply(&mut wb, &op).is_err());
}

#[test]
fn structural_refusals_leave_direct_and_batched_models_unchanged() {
    for op in [
        Op::RenameSheet {
            sheet: SheetId(1),
            name: "Renamed".into(),
        },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(1),
            at: 7,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(1),
            at: 7,
            count: 1,
        },
    ] {
        let mut wb = wb();
        wb.sheets.push(Sheet::new("Inputs"));
        apply(
            &mut wb,
            &add("B2:D5", "expression", None, &["Inputs!$H$12>0"]),
        )
        .unwrap();
        let before = wb.clone();
        let error = apply(&mut wb, &op).unwrap_err();
        assert!(matches!(
            &error,
            xlsx_ops::OpError::InvalidConditionalFormat(_)
        ));
        assert!(error.to_string().contains("Sheet1"));
        assert!(error.to_string().contains("B2:D5"));
        assert!(error.to_string().contains("cf1"));
        assert_eq!(wb, before);
        assert!(
            xlsx_ops::apply_ops(
                &mut wb,
                &[
                    Op::RenameSheet {
                        sheet: SheetId(0),
                        name: "Changed".into()
                    },
                    op
                ]
            )
            .is_err()
        );
        assert_eq!(wb, before);
    }
}
