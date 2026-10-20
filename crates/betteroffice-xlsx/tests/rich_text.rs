use betteroffice_xlsx::{
    CalculationOptions, CellRef, CellState, CellValue, Op, Sheet, SheetId, Workbook, WorkbookModel,
};

fn source() -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = br#"<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>plain</t></is></c><c r="D1" t="inlineStr"><is><r><rPr><i/></rPr><t>Inline</t></r></is></c></row></sheetData></worksheet>"#.to_vec();
    let strings = br#"<sst><si><r><rPr><b/></rPr><t>Rich</t></r></si></sst>"#.to_vec();
    match parts
        .iter_mut()
        .find(|(path, _)| path == "xl/sharedStrings.xml")
    {
        Some(part) => part.1 = strings,
        None => parts.push(("xl/sharedStrings.xml".into(), strings)),
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn rich_queries_preserve_no_edit_save_and_plain_cells() {
    let bytes = source();
    let workbook = Workbook::open(&bytes).unwrap();
    let rich = workbook
        .cell_text_runs(SheetId(0), CellRef::new(0, 0))
        .unwrap();
    assert_eq!(rich[0].style.bold, Some(true));
    assert!(std::ptr::eq(
        rich,
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 1))
            .unwrap()
    ));
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 2))
            .is_none()
    );
    assert_eq!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .unwrap()[0]
            .style
            .italic,
        Some(true)
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(99), CellRef::new(0, 0))
            .is_none()
    );
    assert_eq!(workbook.save().unwrap(), bytes);
}

#[test]
fn content_writes_do_not_restore_runs_after_undo_or_editing_back() {
    for col in [0, 3] {
        let mut workbook = Workbook::open(&source()).unwrap();
        let at = CellRef::new(0, col);
        assert!(workbook.cell_text_runs(SheetId(0), at).is_some());
        workbook
            .edit_cell(SheetId(0), at, "replacement", CalculationOptions::default())
            .unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        workbook.undo(CalculationOptions::default()).unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        workbook.redo(CalculationOptions::default()).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                at,
                if col == 0 { "Rich" } else { "Inline" },
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(workbook.cell_text_runs(SheetId(0), at).is_none());
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(0, 1))
                .is_some()
        );
    }
}

#[test]
fn equal_input_noop_keeps_source_runs() {
    let mut workbook = Workbook::open(&source()).unwrap();
    let result = workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "Rich",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!result.applied);
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
}

#[test]
fn committed_equal_content_assignment_invalidates_only_its_source_cell() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(0, 0),
                    cell: CellState {
                        value: CellValue::Text {
                            value: "Rich".into(),
                        },
                        ..CellState::default()
                    },
                },
                Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(1, 0),
                    cell: CellState {
                        value: CellValue::Number { value: 2.0 },
                        ..CellState::default()
                    },
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 1))
            .is_some()
    );
}

#[test]
fn rejected_transaction_does_not_invalidate_source_runs() {
    let mut workbook = Workbook::open(&source()).unwrap();
    let result = workbook.apply_ops(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: CellValue::Text {
                        value: "replacement".into(),
                    },
                    ..CellState::default()
                },
            },
            Op::SetCell {
                sheet: SheetId(99),
                at: CellRef::new(0, 0),
                cell: CellState::default(),
            },
        ],
        CalculationOptions::default(),
    );
    assert!(result.is_err());
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
}

#[test]
fn source_bindings_follow_rows_and_sheets() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    for col in [0, 3] {
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(0, col))
                .is_none()
        );
        assert!(
            workbook
                .cell_text_runs(SheetId(0), CellRef::new(2, col))
                .is_some()
        );
    }
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(2, 3),
            "changed",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}

#[test]
fn collaborative_sessions_return_no_runs() {
    let workbook = Workbook::open_collaborative(&source(), 7).unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}

#[test]
fn style_only_edits_preserve_runs_and_sheet_reordering_keeps_source_identity() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: betteroffice_xlsx::CellRange::new(CellRef::new(0, 0), CellRef::new(0, 3)),
                patch: betteroffice_xlsx::StylePatch {
                    bold: Some(false),
                    ..Default::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_some()
    );
    workbook
        .apply_ops(
            vec![
                Op::AddSheet {
                    index: 1,
                    name: "Other".into(),
                },
                Op::MoveSheet { from: 0, to: 1 },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(1), CellRef::new(0, 0))
            .is_some()
    );
    assert!(
        workbook
            .cell_text_runs(SheetId(1), CellRef::new(0, 3))
            .is_some()
    );
}

#[test]
fn recovery_without_run_provenance_stays_plain() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "changed",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (restored, _, _, _) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert!(
        restored
            .cell_text_runs(SheetId(0), CellRef::new(0, 0))
            .is_none()
    );
    assert!(
        restored
            .cell_text_runs(SheetId(0), CellRef::new(0, 3))
            .is_none()
    );
}
