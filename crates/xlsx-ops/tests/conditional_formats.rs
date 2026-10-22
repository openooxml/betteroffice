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
fn sheet_rename_rewrites_conditional_formulas_and_undo() {
    let mut wb = wb();
    apply(
        &mut wb,
        &add("A1:A5", "expression", None, &["'Sheet1'!$A$1>0"]),
    )
    .unwrap();
    wb.sheets.push(Sheet::new("Other"));
    let mut other = add("B1:B5", "expression", None, &["'Sheet1'!$A$1>0"]);
    if let Op::AddConditionalFormat { sheet, .. } = &mut other {
        *sheet = SheetId(1);
    }
    apply(&mut wb, &other).unwrap();
    let before = wb.clone();
    let inverse = apply(
        &mut wb,
        &Op::RenameSheet {
            sheet: SheetId(0),
            name: "New name".into(),
        },
    )
    .unwrap();
    for sheet in &wb.sheets {
        assert_eq!(
            sheet.conditional_formats[0].rules[0].formulas,
            ["'New name'!$A$1>0"]
        );
    }
    xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
    assert_eq!(wb, before);
}

#[test]
fn grid_edits_remap_conditional_ranges_and_formulas_and_undo() {
    for (op, range, formula) in [
        (
            Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "A3:B6",
            "$B$6>0",
        ),
        (
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "A2:B4",
            "$B$4>0",
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            "B2:C5",
            "$C$5>0",
        ),
        (
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            "A2:A5",
            "$A$5>0",
        ),
    ] {
        let mut wb = wb();
        apply(&mut wb, &add("A2:B5", "expression", None, &["$B$5>0"])).unwrap();
        let before = wb.clone();
        let inverse = apply(&mut wb, &op).unwrap();
        let group = &wb.sheets[0].conditional_formats[0];
        assert_eq!(group.sqref, range);
        assert_eq!(group.rules[0].formulas, [formula]);
        xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb, before);
    }
}

#[test]
fn deleted_conditional_anchors_rebase_relative_and_mixed_references_and_undo() {
    for (range, formula, op, expected_range, expected_formula) in [
        (
            "A2:A5",
            "B2>0",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "A2:A4",
            "B2>0",
        ),
        (
            "A2:A5",
            "$B2+B$5+$B$2>0",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "A2:A4",
            "$B2+B$4+#REF!>0",
        ),
        (
            "B1:E1",
            "B2>0",
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
            "B1:D1",
            "B2>0",
        ),
        (
            "A2:A5",
            "SUM(B2:C2)>0",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            "A2:A3",
            "SUM(B2:C2)>0",
        ),
    ] {
        let mut wb = wb();
        apply(&mut wb, &add(range, "expression", None, &[formula])).unwrap();
        let before = wb.clone();
        let inverse = apply(&mut wb, &op).unwrap();
        let group = &wb.sheets[0].conditional_formats[0];
        assert_eq!(group.sqref, expected_range);
        assert_eq!(group.rules[0].formulas, [expected_formula]);
        xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb, before);
    }
}

#[test]
fn conditional_anchor_rebases_across_deleted_sqref_areas() {
    let mut wb = wb();
    apply(&mut wb, &add("A2", "expression", None, &["B2>0"])).unwrap();
    wb.sheets[0].conditional_formats[0].sqref = "A2 A4:A5".into();
    apply(
        &mut wb,
        &Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    )
    .unwrap();
    let group = &wb.sheets[0].conditional_formats[0];
    assert_eq!(group.sqref, "A3:A4");
    assert_eq!(group.rules[0].formulas, ["B3>0"]);
}

#[test]
fn deleted_conditional_anchor_rebases_references_to_other_sheets() {
    let mut wb = wb();
    wb.sheets.push(Sheet::new("Other"));
    apply(&mut wb, &add("A2:A5", "expression", None, &["Other!B2>0"])).unwrap();
    apply(
        &mut wb,
        &Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    )
    .unwrap();
    let group = &wb.sheets[0].conditional_formats[0];
    assert_eq!(group.sqref, "A2:A4");
    assert_eq!(group.rules[0].formulas, ["Other!B3>0"]);
}

#[test]
fn nonuniform_conditional_reference_rewrites_refuse_atomically() {
    for (range, formula, op) in [
        (
            "A2:A5",
            "B1>0",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        ),
        (
            "A1:A5",
            "B2>0",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        ),
        (
            "A1:A5",
            "B2>0",
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            },
        ),
        (
            "B1:E1",
            "A2>0",
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        ),
    ] {
        let mut wb = wb();
        apply(&mut wb, &add(range, "expression", None, &[formula])).unwrap();
        let before = wb.clone();
        assert!(matches!(
            apply(&mut wb, &op),
            Err(xlsx_ops::OpError::InvalidConditionalFormat(_))
        ));
        assert_eq!(wb, before);
    }
}

#[test]
fn deleted_conditional_ranges_are_restored_by_undo() {
    let mut wb = wb();
    apply(&mut wb, &add("A2:B3", "expression", None, &["B2>0"])).unwrap();
    let before = wb.clone();
    let inverse = apply(
        &mut wb,
        &Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        },
    )
    .unwrap();
    assert!(wb.sheets[0].conditional_formats.is_empty());
    xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
    assert_eq!(wb, before);
}

#[test]
fn inserted_conditional_ranges_clamp_at_sheet_edges_and_undo() {
    for (op, original, expected) in [
        (
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
            "A1:A1048576",
            "A1:A1048576",
        ),
        (
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 2,
            },
            "A5:A1048575",
            "A7:A1048576",
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
            "A1:XFD1",
            "A1:XFD1",
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 4,
                count: 2,
            },
            "E1:XFC1",
            "G1:XFD1",
        ),
    ] {
        let mut wb = wb();
        apply(&mut wb, &add(original, "cellIs", Some("equal"), &["1"])).unwrap();
        let before = wb.clone();
        let inverse = apply(&mut wb, &op).unwrap();
        assert_eq!(wb.sheets[0].conditional_formats[0].sqref, expected);
        xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb, before);
    }
}

#[test]
fn inserted_conditional_ranges_cannot_silently_fall_off_sheet() {
    for (range, op) in [
        (
            "A1048576",
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
        ),
        (
            "XFD1",
            Op::InsertCols {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
        ),
    ] {
        let mut wb = wb();
        apply(&mut wb, &add(range, "cellIs", Some("equal"), &["1"])).unwrap();
        let before = wb.clone();
        assert!(matches!(
            apply(&mut wb, &op),
            Err(xlsx_ops::OpError::InvalidConditionalFormat(_))
        ));
        assert_eq!(wb, before);
    }
}

#[test]
fn unchanged_clamped_conditional_formulas_on_other_sheets_restore_on_undo() {
    let mut wb = wb();
    wb.sheets.push(Sheet::new("Other"));
    apply(
        &mut wb,
        &add("A1", "expression", None, &["SUM(Other!$1:$1048576)>0"]),
    )
    .unwrap();
    let before = wb.clone();
    let inverse = apply(
        &mut wb,
        &Op::InsertRows {
            sheet: SheetId(1),
            at: 4,
            count: 1,
        },
    )
    .unwrap();
    assert_eq!(
        wb.sheets[0].conditional_formats,
        before.sheets[0].conditional_formats
    );
    xlsx_ops::apply_ops(&mut wb, &inverse.0).unwrap();
    assert_eq!(wb, before);
}

#[test]
fn unsupported_conditional_rename_leaves_cell_formulas_and_sheet_unchanged() {
    let mut wb = wb();
    apply(&mut wb, &add("A1", "expression", None, &["'Sheet1'!A1#"])).unwrap();
    wb.sheets[0].set_cell(
        xlsx_model::CellRef::new(0, 0),
        xlsx_model::Cell {
            formula: Some("Sheet1!B1".into()),
            ..Default::default()
        },
    );
    let before = wb.clone();
    assert!(matches!(
        apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into()
            }
        ),
        Err(xlsx_ops::OpError::InvalidConditionalFormat(_))
    ));
    assert_eq!(wb, before);
}
