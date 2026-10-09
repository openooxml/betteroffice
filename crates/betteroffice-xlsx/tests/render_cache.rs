use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellValue, EditRequest, FreezePane,
    Op, ProposalEditInput, ProposalRequest, Sheet, SheetId, StylePatch, Viewport, Workbook,
    WorkbookModel,
};
use xlsx_render::geometry::{geometry_counters, reset_geometry_counters};

fn model(rows: u32, cols: u32) -> WorkbookModel {
    let mut sheet = Sheet::new("Data");
    sheet.format.default_row_height_pt = Some(15.0);
    sheet.freeze_pane = Some(FreezePane {
        rows: 1,
        cols: 1,
        top_left: CellRef::new(1, 1),
    });
    sheet.col_widths.insert(2, 18.0);
    sheet.row_heights.insert(3, 24.0);
    for row in 0..rows {
        for col in 0..cols {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row * cols + col),
                    },
                    ..Cell::default()
                },
            );
        }
    }
    WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    }
}

fn viewport(index: u32) -> Viewport {
    Viewport {
        x: (index % 5) as f32 * 70.0,
        y: index as f32 * 140.0,
        width: 640.0,
        height: 400.0,
    }
}

fn frame(workbook: &Workbook, sheet: SheetId, viewport: &Viewport) -> Vec<u8> {
    serde_json::to_vec(&workbook.display_list_for(sheet, viewport).unwrap()).unwrap()
}

fn assert_rebuilt(workbook: &Workbook) {
    let sheet = workbook.active_sheet();
    reset_geometry_counters();
    let actual = frame(workbook, sheet, &viewport(0));
    assert_eq!(geometry_counters().0, 1);
    assert_eq!(
        geometry_counters().1,
        workbook.sheet(sheet).unwrap().iter_cells().count() as u64,
    );
    reset_geometry_counters();
    assert_eq!(frame(workbook, sheet, &viewport(0)), actual);
    let info = workbook.sheet_info().unwrap();
    let scroll = workbook
        .cell_scroll_position(sheet, CellRef::new(12, 4))
        .unwrap();
    assert!(
        workbook
            .chart_at_point(&viewport(0), 20.0, 20.0)
            .unwrap()
            .is_none()
    );
    assert_eq!(geometry_counters(), (0, 0));

    let mut cold = Workbook::from_model(workbook.model().clone()).unwrap();
    cold.set_active_sheet(sheet).unwrap();
    assert_eq!(actual, frame(&cold, sheet, &viewport(0)));
    assert_eq!(
        actual,
        serde_json::to_vec(
            &xlsx_render::build_display_list(cold.model(), sheet, &viewport(0)).unwrap(),
        )
        .unwrap(),
    );
    assert_eq!(info, cold.sheet_info().unwrap());
    assert_eq!(
        scroll,
        cold.cell_scroll_position(sheet, CellRef::new(12, 4))
            .unwrap(),
    );
}

#[test]
fn warm_scroll_does_no_geometry_or_autofit_work() {
    let workbook = Workbook::from_model(model(1_000, 20)).unwrap();
    reset_geometry_counters();
    workbook.display_list(&viewport(0)).unwrap();
    assert_eq!(geometry_counters(), (1, 20_000));

    reset_geometry_counters();
    let frames: Vec<_> = (0..30)
        .map(|index| {
            let viewport = viewport(index);
            workbook.sheet_info().unwrap();
            workbook
                .cell_scroll_position(SheetId(0), CellRef::new(index * 7, 4))
                .unwrap();
            workbook.chart_at_point(&viewport, 50.0, 50.0).unwrap();
            serde_json::to_vec(&workbook.display_list(&viewport).unwrap()).unwrap()
        })
        .collect();
    assert_eq!(geometry_counters(), (0, 0));

    let cold = Workbook::from_model(workbook.model().clone()).unwrap();
    for (index, actual) in frames.iter().enumerate() {
        let viewport = viewport(index as u32);
        assert_eq!(*actual, frame(&cold, SheetId(0), &viewport));
        assert_eq!(
            *actual,
            serde_json::to_vec(
                &xlsx_render::build_display_list(cold.model(), SheetId(0), &viewport).unwrap(),
            )
            .unwrap(),
        );
    }
}

#[test]
fn edits_dimensions_fonts_merges_and_history_invalidate_geometry() {
    let mut workbook = Workbook::from_model(model(20, 5)).unwrap();
    let options = CalculationOptions::default();
    assert_rebuilt(&workbook);

    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 0), "42", options)
        .unwrap();
    assert_rebuilt(&workbook);

    for op in [
        Op::SetRowHeight {
            sheet: SheetId(0),
            row: 2,
            height: Some(36.0),
        },
        Op::SetColWidth {
            sheet: SheetId(0),
            col: 1,
            width: Some(22.0),
        },
        Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::new(CellRef::new(4, 0), CellRef::new(4, 1)),
            patch: StylePatch {
                font_size: Some(28.0),
                ..StylePatch::default()
            },
        },
        Op::MergeCells {
            sheet: SheetId(0),
            range: CellRange::new(CellRef::new(4, 0), CellRef::new(5, 1)),
        },
        Op::UnmergeCells {
            sheet: SheetId(0),
            range: CellRange::new(CellRef::new(4, 0), CellRef::new(5, 1)),
        },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::SetFreezePane {
            sheet: SheetId(0),
            pane: None,
        },
    ] {
        assert!(workbook.apply_ops(vec![op], options).unwrap().applied);
        assert_rebuilt(&workbook);
    }
    assert!(workbook.undo(options).unwrap().applied);
    assert_rebuilt(&workbook);
    assert!(workbook.redo(options).unwrap().applied);
    assert_rebuilt(&workbook);
    workbook.recalculate_all(options);
    assert_rebuilt(&workbook);
}

#[test]
fn batch_and_proposal_commits_invalidate_geometry() {
    let mut workbook = Workbook::from_model(model(20, 5)).unwrap();
    let options = CalculationOptions::default();
    assert_rebuilt(&workbook);
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(0, 0),
                input: "100".into(),
            }],
            options,
        )
        .unwrap();
    assert_rebuilt(&workbook);

    let request: EditRequest = serde_json::from_value(serde_json::json!({
        "expectVersion": workbook.version(),
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A1" } },
            "inputs": [["200"]],
        }],
    }))
    .unwrap();
    workbook.apply_edits(&request).unwrap().unwrap();
    assert_rebuilt(&workbook);

    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "review".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: CellRef::new(0, 0),
                    input: "300".into(),
                    number_format: None,
                }],
            },
            options,
        )
        .unwrap();
    workbook
        .accept_proposal(&proposal.id, false, options)
        .unwrap();
    assert_rebuilt(&workbook);
}

#[test]
fn remote_history_and_snapshot_adoption_invalidate_geometry() {
    let model = model(20, 5);
    let options = CalculationOptions::default();
    let mut local = Workbook::from_model_collaborative(model.clone(), 101).unwrap();
    let mut peer = Workbook::from_model_collaborative(model.clone(), 102).unwrap();
    assert_rebuilt(&local);
    local
        .edit_cell(SheetId(0), CellRef::new(0, 0), "42", options)
        .unwrap();
    assert_rebuilt(&local);
    assert!(local.undo(options).unwrap().applied);
    assert_rebuilt(&local);
    assert!(local.redo(options).unwrap().applied);
    assert_rebuilt(&local);

    let baseline = local.encode_state_vector_v1();
    peer.apply_update_v1(&local.encode_state_as_update_v1(), options)
        .unwrap();
    peer.patch_range_style(
        SheetId(0),
        CellRange::new(CellRef::new(6, 0), CellRef::new(6, 0)),
        StylePatch {
            font_size: Some(32.0),
            ..StylePatch::default()
        },
        options,
    )
    .unwrap();
    let update = peer.encode_diff_v1(&baseline).unwrap();
    assert!(local.apply_update_v1(&update, options).unwrap().applied);
    assert_rebuilt(&local);
    assert_eq!(local.model(), peer.model());

    let mut restored = Workbook::from_model_collaborative(model, 103).unwrap();
    assert_rebuilt(&restored);
    assert!(
        restored
            .apply_update_v1(&peer.encode_state_as_update_v1(), options)
            .unwrap()
            .applied
    );
    assert_rebuilt(&restored);
    assert_eq!(restored.model(), peer.model());
}

#[test]
fn sheet_info_shares_geometry_and_keeps_bounds_folding() {
    let mut model = model(20, 5);
    model.sheets[0].format.default_row_height_pt = None;
    let mut workbook = Workbook::from_model(model).unwrap();
    reset_geometry_counters();
    workbook.sheet_info().unwrap();
    assert_eq!(geometry_counters(), (1, 100));
    workbook.display_list(&viewport(0)).unwrap();
    assert_eq!(geometry_counters(), (1, 100));

    let at = CellRef::new(22, 6);
    workbook
        .edit_cell(SheetId(0), at, "1", CalculationOptions::default())
        .unwrap();
    reset_geometry_counters();
    let info = workbook.sheet_info().unwrap();
    assert_eq!(geometry_counters(), (0, 0));
    let cold = Workbook::from_model(workbook.model().clone()).unwrap();
    assert_eq!(info, cold.sheet_info().unwrap());
    assert_rebuilt(&workbook);

    workbook
        .edit_cell(SheetId(0), at, "", CalculationOptions::default())
        .unwrap();
    assert_rebuilt(&workbook);
}

#[test]
fn same_value_edit_keeps_warm_geometry() {
    let mut workbook = Workbook::from_model(model(20, 5)).unwrap();
    let before = frame(&workbook, SheetId(0), &viewport(0));
    reset_geometry_counters();
    assert!(
        !workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "0",
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
    assert_eq!(before, frame(&workbook, SheetId(0), &viewport(0)));
    assert_eq!(geometry_counters(), (0, 0));
}

#[test]
fn switching_sheets_reuses_each_sheets_geometry() {
    let mut model = model(20, 5);
    let mut second = model.sheets[0].clone();
    second.name = "Other".into();
    second.row_heights.insert(5, 40.0);
    model.sheets.push(second);
    let mut workbook = Workbook::from_model(model).unwrap();
    reset_geometry_counters();
    frame(&workbook, SheetId(0), &viewport(0));
    assert_eq!(geometry_counters(), (1, 100));
    frame(&workbook, SheetId(1), &viewport(0));
    assert_eq!(geometry_counters(), (2, 200));

    reset_geometry_counters();
    for sheet in [SheetId(1), SheetId(0), SheetId(1), SheetId(0)] {
        workbook.set_active_sheet(sheet).unwrap();
        workbook.sheet_info().unwrap();
        workbook.display_list(&viewport(1)).unwrap();
    }
    assert_eq!(geometry_counters(), (0, 0));
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(0, 0),
            "77",
            CalculationOptions::default(),
        )
        .unwrap();
    for sheet in [SheetId(0), SheetId(1)] {
        workbook.set_active_sheet(sheet).unwrap();
        assert_rebuilt(&workbook);
    }
}
