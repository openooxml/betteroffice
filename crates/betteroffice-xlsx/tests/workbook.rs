#[cfg(feature = "raster")]
use betteroffice_xlsx::RenderOptions;
use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellState, CellValue, DrawCmd, Error,
    MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID, MAX_COLLABORATION_STATE_VECTOR_ENTRIES,
    NumberFormatKind, NumberFormatMutation, Op, ProposalEditInput, ProposalRequest, Sheet, SheetId,
    StylePatch, UpdateOrigin, Viewport, Workbook, WorkbookModel,
};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use yrs::Update as YrsUpdate;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn sample_parts() -> Vec<(String, Vec<u8>)> {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Number { value: 10.0 },
            style: Some(0),
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("A2"),
        Cell {
            value: CellValue::Number { value: 5.0 },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("B1"),
        Cell {
            value: CellValue::Number { value: 999.0 },
            formula: Some("SUM(A1:A2)".into()),
            ..Cell::default()
        },
    );
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Default::default());
    model.sheets.push(sheet);
    model.sheets.push(Sheet::new("Empty"));
    xlsx_parse::serialize_workbook(&model).unwrap()
}

fn sample_xlsx() -> Vec<u8> {
    ooxml_opc::rezip_parts(&sample_parts()).unwrap()
}

fn overlapping_merge_parts() -> Vec<(String, Vec<u8>)> {
    let workbook =
        r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels = r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let worksheet = r#"<worksheet><sheetData/><mergeCells count="5"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="C3:D4"/><mergeCell ref="D4:E5"/><mergeCell ref="F1:G1"/></mergeCells></worksheet>"#;
    vec![
        ("xl/workbook.xml".to_string(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels.as_bytes().to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            worksheet.as_bytes().to_vec(),
        ),
    ]
}

#[test]
fn open_and_recalculation_are_explicit() {
    let cached = Workbook::open(&sample_xlsx()).unwrap();
    assert_eq!(
        cached
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 999.0 }
    );

    let calculated =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    assert_eq!(
        calculated
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 15.0 }
    );

    let mut read_only = Workbook::open_for_read(&sample_xlsx()).unwrap();
    let result = read_only
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert_eq!(result.changed[0].cell, cell("B1"));
}

#[test]
fn edits_recalculate_render_and_round_trip() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    let result = workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert_eq!(result.changed.len(), 1);
    assert_eq!(result.changed[0].cell, cell("B1"));
    assert_eq!(
        workbook
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("A1"))
            .unwrap()
            .style,
        Some(0)
    );

    let display = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 240.0,
            height: 120.0,
        })
        .unwrap();
    assert!(
        display
            .commands
            .iter()
            .any(|command| { matches!(command, DrawCmd::Text { text, .. } if text == "25") })
    );

    #[cfg(feature = "raster")]
    {
        let png = workbook
            .render_sheet(
                SheetId(0),
                &RenderOptions {
                    range: Some(betteroffice_xlsx::CellRange::parse_a1("A1:B2").unwrap()),
                    ..RenderOptions::default()
                },
            )
            .unwrap();
        assert_eq!(
            &png.bytes[..8],
            &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]
        );
    }

    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.cell(SheetId(0), cell("A1")).unwrap().input, "20");
    assert_eq!(
        reopened.cell(SheetId(0), cell("B1")).unwrap().input,
        "=SUM(A1:A2)"
    );
    assert_eq!(
        reopened
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 25.0 }
    );
}

#[test]
fn yrs_state_tracks_structural_edits_and_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::AddSheet {
                    index: 1,
                    name: "Inserted".into(),
                },
                Op::SetCell {
                    sheet: SheetId(1),
                    at: cell("C3"),
                    cell: CellState {
                        value: CellValue::Text {
                            value: "shared".into(),
                        },
                        ..CellState::default()
                    },
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.sheet_count(), 3);
    assert_eq!(reopened.sheet_id("Inserted"), Some(SheetId(1)));
    assert_eq!(
        reopened.cell(SheetId(1), cell("C3")).unwrap().input,
        "shared"
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.sheet_count(), 2);
    assert_eq!(reopened.sheet_id("Inserted"), None);

    workbook.redo(CalculationOptions::default()).unwrap();
    let model = workbook.into_model();
    assert_eq!(model.sheets.len(), 3);
    assert_eq!(model.sheets[1].name, "Inserted");
    assert_eq!(
        model.sheets[1].cell(cell("C3")).unwrap().value,
        CellValue::Text {
            value: "shared".into()
        }
    );
}

#[test]
fn standalone_removed_sheet_state_encodes_and_undo_restores_the_model() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    let original = workbook.model().clone();
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();

    assert_eq!(workbook.sheet_count(), 1);
    assert!(!workbook.encode_state_as_update_v1().is_empty());
    assert_eq!(
        Workbook::open(&workbook.save().unwrap())
            .unwrap()
            .sheet_count(),
        1
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model(), &original);
    assert!(!workbook.encode_state_vector_v1().is_empty());
    assert!(!workbook.encode_state_as_update_v1().is_empty());
}

#[test]
fn workbook_remains_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Workbook>();
}

#[test]
fn undo_redo_and_proposals_share_the_typed_session() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert!(workbook.can_undo());
    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "10");
    assert!(
        workbook
            .redo(CalculationOptions::default())
            .unwrap()
            .applied
    );

    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: Some("update total".into()),
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "30".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(proposal.id, "p1");
    assert_eq!(workbook.proposals().len(), 1);
    let accepted = workbook
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert_eq!(accepted.proposal_id, "p1");
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "30");
    assert!(workbook.proposals().is_empty());
}

#[test]
fn rejects_empty_workbook_ops_atomically() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let result = workbook.apply_ops(
        vec![Op::RemoveSheet { index: 1 }, Op::RemoveSheet { index: 0 }],
        CalculationOptions::default(),
    );
    assert!(matches!(result, Err(Error::NoSheets)));
    assert_eq!(workbook.sheet_count(), 2);
}

#[test]
fn rejects_overlapping_merged_ranges() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    sheet.merges = vec![
        CellRange::parse_a1("A1:B2").unwrap(),
        CellRange::parse_a1("B2:C3").unwrap(),
    ];
    model.sheets.push(sheet);

    assert!(matches!(
        Workbook::from_model(model),
        Err(Error::InvalidOperation(message))
            if message == "workbook contains overlapping merged ranges"
    ));
}

#[test]
fn parsed_overlapping_merges_open_and_save_normalized() {
    let model = xlsx_parse::parse_workbook(&overlapping_merge_parts()).unwrap();
    let merges: Vec<_> = model.sheets[0]
        .merges
        .iter()
        .map(|range| range.to_a1())
        .collect();
    assert_eq!(merges, ["A1:B2", "C3:D4", "F1:G1"]);

    let workbook = Workbook::from_model(model).unwrap();
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet_xml = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .map(|(_, bytes)| std::str::from_utf8(bytes).unwrap())
        .unwrap();
    assert!(sheet_xml.contains(
        r#"<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="C3:D4"/><mergeCell ref="F1:G1"/></mergeCells>"#
    ));

    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].merges,
        workbook.model().sheets[0].merges
    );
}

#[test]
fn validates_raw_ops_and_noop_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let result = workbook.edit_cells(
        SheetId(0),
        &Vec::<CellInput>::new(),
        CalculationOptions::default(),
    );
    assert!(!result.unwrap().applied);
    assert!(!workbook.can_undo());

    let invalid = workbook.apply_ops(
        vec![Op::SetColWidth {
            sheet: SheetId(0),
            col: 1_000_000_000,
            width: Some(12.0),
        }],
        CalculationOptions::default(),
    );
    assert!(matches!(invalid, Err(Error::InvalidOperation(_))));
    assert!(!workbook.can_undo());

    let duplicate_name = workbook.apply_ops(
        vec![Op::RenameSheet {
            sheet: SheetId(0),
            name: "Empty".into(),
        }],
        CalculationOptions::default(),
    );
    assert!(matches!(duplicate_name, Err(Error::InvalidOperation(_))));

    let shifted_dimension = workbook.apply_ops(
        vec![
            Op::SetRowHeight {
                sheet: SheetId(0),
                row: betteroffice_xlsx::MAX_ROWS - 1,
                height: Some(20.0),
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: betteroffice_xlsx::MAX_ROWS,
            },
        ],
        CalculationOptions::default(),
    );
    assert!(matches!(shifted_dimension, Err(Error::InvalidOperation(_))));
    assert!(!workbook.can_undo());
}

#[test]
fn semantic_noop_preserves_redo_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(workbook.can_redo());

    let formula_result = workbook
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: cell("B1"),
                    input: "=1".into(),
                },
                CellInput {
                    cell: cell("B1"),
                    input: "=SUM(A1:A2)".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!formula_result.applied);
    assert!(workbook.can_redo());

    let result = workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    assert!(!result.applied);
    assert!(workbook.can_redo());
}

#[test]
fn rejects_insertions_that_discard_boundary_content() {
    let mut sheet = Sheet::new("Data");
    let last_row = CellRef::new(betteroffice_xlsx::MAX_ROWS - 1, 0);
    let last_col = CellRef::new(0, betteroffice_xlsx::MAX_COLS - 1);
    sheet.set_cell(
        last_row,
        Cell {
            value: CellValue::Text {
                value: "row edge".into(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        last_col,
        Cell {
            value: CellValue::Text {
                value: "column edge".into(),
            },
            ..Cell::default()
        },
    );
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    let mut workbook = Workbook::from_model(model).unwrap();

    let row_error = workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(row_error, Error::InvalidOperation(_)));
    assert_eq!(
        workbook.cell(SheetId(0), last_row).unwrap().input,
        "row edge"
    );

    let column_error = workbook
        .apply_ops(
            vec![Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(column_error, Error::InvalidOperation(_)));
    assert_eq!(
        workbook.cell(SheetId(0), last_col).unwrap().input,
        "column edge"
    );
    assert!(!workbook.can_undo());
}

#[test]
fn rejects_reversed_ranges_and_oversized_dimensions() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let reversed = CellRange {
        start: cell("B2"),
        end: cell("A1"),
    };
    assert!(matches!(
        workbook.range_cells(SheetId(0), reversed),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::MergeCells {
                sheet: SheetId(0),
                range: reversed,
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::SetColWidth {
                sheet: SheetId(0),
                col: 0,
                width: Some(256.0),
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::SetRowHeight {
                sheet: SheetId(0),
                row: 0,
                height: Some(410.0),
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.edit_cell(
            SheetId(0),
            cell("A1"),
            &"x".repeat(xlsx_calc::eval::MAX_CELL_TEXT_CHARS + 1),
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
}

#[test]
fn proposal_staleness_uses_cell_state_not_display_text() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("B1"),
                    input: "1".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("B1"),
            "=999",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(matches!(
        workbook.accept_proposal("p1", false, CalculationOptions::default()),
        Err(Error::StaleProposal(_))
    ));
}

#[test]
fn proposal_acceptance_applies_duplicate_targets_sequentially() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A1"),
                        input: "20".into(),
                    },
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A1"),
                        input: "30".into(),
                    },
                ],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "30");
}

#[test]
fn rename_invalidates_pending_proposals() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "=Data!A2".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.proposals().is_empty());
}

#[test]
fn reports_recalculation_limits_without_overwriting_cached_values() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let mut formulas = Sheet::new("Formulas");
    formulas.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Number { value: 123.0 },
            formula: Some("SUM(Data!A1:XFD1048576)".into()),
            ..Cell::default()
        },
    );
    model.sheets.push(formulas);
    let bytes = ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap();
    let workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(
        workbook.model().sheets[1].cell(cell("A1")).unwrap().value,
        CellValue::Number { value: 123.0 }
    );
    assert_eq!(workbook.last_calculation().limited_cells.len(), 1);
}

#[test]
fn structural_ops_invalidate_coordinate_proposals() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "30".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.proposals().is_empty());
}

#[test]
fn display_lists_do_not_inherit_raster_dimension_caps() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    assert!(
        workbook
            .display_list(&Viewport {
                x: 0.0,
                y: 0.0,
                width: 20_000.0,
                height: 120.0,
            })
            .is_ok()
    );
}

#[test]
fn display_lists_reject_excessive_cell_spans() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    let error = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 6_000_000.0,
        })
        .unwrap_err();
    assert!(matches!(error, Error::DisplayTooLarge { .. }));
}

#[cfg(feature = "raster")]
#[test]
fn raster_rejects_excessive_total_pixel_area() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    let error = workbook
        .render_png(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 5_000.0,
            height: 5_000.0,
        })
        .unwrap_err();
    assert!(matches!(error, Error::RenderAreaTooLarge { .. }));
}

#[test]
fn collaboration_vectors_diffs_and_deterministic_baseline_handshake() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 101).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 202).unwrap();

    assert_eq!(left.client_id(), 101);
    assert_eq!(right.client_id(), 202);
    assert_ne!(left.client_id(), right.client_id());
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_eq!(
        left.encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap(),
        &[0, 0]
    );

    left.edit_cell(SheetId(0), cell("A1"), "21", CalculationOptions::default())
        .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    assert!(
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(left.model(), right.model());
    assert!(
        !left
            .apply_update_v1(
                &right
                    .encode_diff_v1(&left.encode_state_vector_v1())
                    .unwrap(),
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
}

#[test]
fn duplicate_runtime_client_ids_are_an_invalid_host_configuration() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 211).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 211).unwrap();
    let baseline = left.encode_state_vector_v1();

    left.edit_cell(
        SheetId(0),
        cell("C1"),
        "left",
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .edit_cell(
            SheetId(0),
            cell("C2"),
            "right",
            CalculationOptions::default(),
        )
        .unwrap();
    let from_left = left.encode_diff_v1(&baseline).unwrap();
    let from_right = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();

    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_ne!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_ne!(left.model(), right.model());
}

#[test]
fn collaborative_undo_redo_track_only_local_user_edits() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 221).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert!(workbook.can_undo());
    assert!(!workbook.can_redo());
    assert_eq!(workbook.history_state().undo_depth, 1);

    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "10");
    assert!(!workbook.can_undo());
    assert!(workbook.can_redo());
    assert!(
        workbook
            .redo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "20");

    let mut agent_only = Workbook::open_collaborative(&bytes, 222).unwrap();
    agent_only
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A2"),
                    input: "30".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    agent_only
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert!(!agent_only.can_undo());
    assert!(
        !agent_only
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
}

#[test]
fn collaborative_undo_converges_after_a_concurrent_remote_edit() {
    let bytes = sample_xlsx();
    for (left_id, right_id) in [(231, 230), (230, 231)] {
        let mut left = Workbook::open_collaborative(&bytes, left_id).unwrap();
        let mut right = Workbook::open_collaborative(&bytes, right_id).unwrap();
        let baseline = left.encode_state_vector_v1();

        left.edit_cell(
            SheetId(0),
            cell("C1"),
            "left",
            CalculationOptions::default(),
        )
        .unwrap();
        right
            .edit_cell(
                SheetId(0),
                cell("C1"),
                "right",
                CalculationOptions::default(),
            )
            .unwrap();
        let from_left = left.encode_diff_v1(&baseline).unwrap();
        let from_right = right.encode_diff_v1(&baseline).unwrap();
        left.apply_update_v1(&from_right, CalculationOptions::default())
            .unwrap();
        right
            .apply_update_v1(&from_left, CalculationOptions::default())
            .unwrap();

        let right_before_undo = right.encode_state_vector_v1();
        left.undo(CalculationOptions::default()).unwrap();
        let undo = left.encode_diff_v1(&right_before_undo).unwrap();
        right
            .apply_update_v1(&undo, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert!(right.can_undo());
    }
}

#[test]
fn concurrent_disjoint_and_same_cell_edits_converge() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 301).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 302).unwrap();
    let baseline = left.encode_state_vector_v1();

    left.edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    right
        .edit_cell(SheetId(0), cell("A2"), "7", CalculationOptions::default())
        .unwrap();
    let from_left = left.encode_diff_v1(&baseline).unwrap();
    let from_right = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model(), right.model());
    assert_eq!(left.cell(SheetId(0), cell("A1")).unwrap().input, "20");
    assert_eq!(left.cell(SheetId(0), cell("A2")).unwrap().input, "7");

    let left_before = left.encode_state_vector_v1();
    let right_before = right.encode_state_vector_v1();
    left.edit_cell(
        SheetId(0),
        cell("C1"),
        "left",
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .edit_cell(
            SheetId(0),
            cell("C1"),
            "right",
            CalculationOptions::default(),
        )
        .unwrap();
    let from_left = left.encode_diff_v1(&right_before).unwrap();
    let from_right = right.encode_diff_v1(&left_before).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model(), right.model());
    assert!(matches!(
        left.cell(SheetId(0), cell("C1")).unwrap().input.as_str(),
        "left" | "right"
    ));
}

#[test]
fn concurrent_style_and_content_changes_compose() {
    let bytes = sample_xlsx();
    let mut content = Workbook::open_collaborative(&bytes, 401).unwrap();
    let mut style = Workbook::open_collaborative(&bytes, 402).unwrap();
    let baseline = content.encode_state_vector_v1();

    content
        .edit_cell(SheetId(0), cell("A1"), "25", CalculationOptions::default())
        .unwrap();
    style
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A1"), cell("A1")),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let content_update = content.encode_diff_v1(&baseline).unwrap();
    let style_update = style.encode_diff_v1(&baseline).unwrap();
    content
        .apply_update_v1(&style_update, CalculationOptions::default())
        .unwrap();
    style
        .apply_update_v1(&content_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(content.model(), style.model());
    let composed = content
        .model()
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("A1"))
        .unwrap();
    assert_eq!(composed.value, CellValue::Number { value: 25.0 });
    assert_eq!(
        content
            .selection_formatting(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
            .unwrap()
            .bold,
        Some(true)
    );
}

#[test]
fn collaborative_formatting_round_trips_and_matches_aggregation() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 403).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 404).unwrap();
    let range = CellRange::new(cell("A1"), cell("B2"));

    left.patch_range_style(
        SheetId(0),
        range,
        StylePatch {
            bold: Some(true),
            fill_color: Some("#ffcc00".into()),
            text_color: Some("#123456".into()),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    left.set_range_number_format(
        SheetId(0),
        range,
        NumberFormatMutation::Custom {
            pattern: "0.0000".into(),
        },
        CalculationOptions::default(),
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    assert_eq!(
        left.selection_formatting(SheetId(0), range).unwrap(),
        right.selection_formatting(SheetId(0), range).unwrap()
    );
    let formatting = right.selection_formatting(SheetId(0), range).unwrap();
    assert_eq!(formatting.bold, Some(true));
    assert_eq!(formatting.fill_color.as_deref(), Some("#ffcc00"));
    assert_eq!(formatting.text_color.as_deref(), Some("#123456"));
    assert_eq!(formatting.number_format, Some(NumberFormatKind::Custom));
    assert_eq!(formatting.number_format_pattern.as_deref(), Some("0.0000"));
}

#[test]
fn concurrent_formatting_restyles_converge_with_all_formats_available() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 405).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 406).unwrap();
    let baseline = left.encode_state_vector_v1();
    let range = CellRange::new(cell("A1"), cell("B2"));

    left.patch_range_style(
        SheetId(0),
        range,
        StylePatch {
            bold: Some(true),
            text_color: Some("#aa0000".into()),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            range,
            StylePatch {
                italic: Some(true),
                fill_color: Some("#00aa00".into()),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left.encode_diff_v1(&baseline).unwrap();
    let right_update = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    assert_eq!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_eq!(left.model().styles.cell_xfs.len(), 3);
    assert_eq!(
        left.selection_formatting(SheetId(0), range).unwrap(),
        right.selection_formatting(SheetId(0), range).unwrap()
    );
}

#[test]
fn concurrent_identical_formats_are_content_deduplicated() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 407).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 408).unwrap();
    let baseline = left.encode_state_vector_v1();
    let patch = StylePatch {
        bold: Some(true),
        font_family: Some("Inter".into()),
        ..StylePatch::default()
    };

    left.patch_range_style(
        SheetId(0),
        CellRange::new(cell("A1"), cell("A1")),
        patch.clone(),
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A2"), cell("A2")),
            patch,
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left.encode_diff_v1(&baseline).unwrap();
    let right_update = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    let sheet = &left.model().sheets[0];
    assert_eq!(
        sheet.cell(cell("A1")).unwrap().style,
        sheet.cell(cell("A2")).unwrap().style
    );
    assert_eq!(left.model().styles.cell_xfs.len(), 2);
}

#[test]
fn collaborative_formatting_undo_is_local_origin_only() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 409).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 410).unwrap();

    left.patch_range_style(
        SheetId(0),
        CellRange::new(cell("A1"), cell("A1")),
        StylePatch {
            bold: Some(true),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A2"), cell("A2")),
            StylePatch {
                fill_color: Some("#abcdef".into()),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let right_update = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();
    let format = right
        .capture_format(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
        .unwrap();
    right
        .apply_format(
            SheetId(0),
            CellRange::new(cell("A3"), cell("A3")),
            format,
            CalculationOptions::default(),
        )
        .unwrap();
    let reused_format = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    left.apply_update_v1(&reused_format, CalculationOptions::default())
        .unwrap();
    let right_before_undo = right.encode_state_vector_v1();

    assert!(left.undo(CalculationOptions::default()).unwrap().applied);
    let undo = left.encode_diff_v1(&right_before_undo).unwrap();
    right
        .apply_update_v1(&undo, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    let a1 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
        .unwrap();
    let a2 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A2"), cell("A2")))
        .unwrap();
    let a3 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A3"), cell("A3")))
        .unwrap();
    assert_eq!(a1.bold, Some(false));
    assert_eq!(a2.fill_color.as_deref(), Some("#abcdef"));
    assert_eq!(a3.bold, Some(true));
}

#[test]
fn style_edits_do_not_publish_recalculated_formula_caches_as_content() {
    let bytes = sample_xlsx();
    for (formula_client, style_client) in [(411, 412), (422, 421)] {
        let mut formula = Workbook::open_collaborative_recalculated(
            &bytes,
            formula_client,
            CalculationOptions::default(),
        )
        .unwrap();
        let mut style = Workbook::open_collaborative_recalculated(
            &bytes,
            style_client,
            CalculationOptions::default(),
        )
        .unwrap();
        let baseline = formula.encode_state_vector_v1();

        formula
            .edit_cell(
                SheetId(0),
                cell("B1"),
                "=SUM(A1:A2)+1",
                CalculationOptions::default(),
            )
            .unwrap();
        style
            .apply_ops(
                vec![Op::SetCell {
                    sheet: SheetId(0),
                    at: cell("B1"),
                    cell: CellState {
                        value: CellValue::Number { value: 15.0 },
                        formula: Some("SUM(A1:A2)".into()),
                        style: Some(0),
                    },
                }],
                CalculationOptions::default(),
            )
            .unwrap();

        let formula_update = formula.encode_diff_v1(&baseline).unwrap();
        let style_update = style.encode_diff_v1(&baseline).unwrap();
        formula
            .apply_update_v1(&style_update, CalculationOptions::default())
            .unwrap();
        style
            .apply_update_v1(&formula_update, CalculationOptions::default())
            .unwrap();

        assert_eq!(formula.model(), style.model());
        let composed = formula.model().sheets[0].cell(cell("B1")).unwrap();
        assert_eq!(composed.formula.as_deref(), Some("SUM(A1:A2)+1"));
        assert_eq!(composed.style, Some(0));
    }
}

#[test]
fn remote_formulas_recalculate_locally_and_save_current_caches() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 501).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 502).unwrap();

    left.edit_cell(
        SheetId(0),
        cell("B1"),
        "=A1*2",
        CalculationOptions::default(),
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(right.cell(SheetId(0), cell("B1")).unwrap().input, "=A1*2");
    assert_eq!(
        right.model().sheets[0].cell(cell("B1")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );

    let shared_before_recalc = right.encode_state_as_update_v1();
    right.recalculate_all(CalculationOptions::default());
    assert_eq!(right.encode_state_as_update_v1(), shared_before_recalc);
    let reopened = Workbook::open(&right.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(cell("B1")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );
}

#[test]
fn remote_changed_cells_compare_against_the_current_projection() {
    let bytes = sample_xlsx();
    let options = CalculationOptions::default();
    let mut left = Workbook::open_collaborative_recalculated(&bytes, 511, options).unwrap();
    let mut right = Workbook::open_collaborative_recalculated(&bytes, 512, options).unwrap();

    left.edit_cell(SheetId(0), cell("A1"), "20", options)
        .unwrap();
    let first = right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
            options,
        )
        .unwrap();
    assert_eq!(
        first.changed,
        [
            betteroffice_xlsx::CellAddress {
                sheet: SheetId(0),
                cell: cell("A1"),
            },
            betteroffice_xlsx::CellAddress {
                sheet: SheetId(0),
                cell: cell("B1"),
            },
        ]
    );

    left.edit_cell(SheetId(1), cell("A1"), "unrelated", options)
        .unwrap();
    let second = right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
            options,
        )
        .unwrap();
    assert_eq!(
        second.changed,
        [betteroffice_xlsx::CellAddress {
            sheet: SheetId(1),
            cell: cell("A1"),
        }]
    );
}

#[test]
fn duplicate_and_reversed_update_delivery_are_safe() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open_collaborative(&bytes, 601).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 602).unwrap();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    source
        .edit_cell(SheetId(0), cell("A1"), "31", CalculationOptions::default())
        .unwrap();
    source
        .edit_cell(SheetId(0), cell("A2"), "9", CalculationOptions::default())
        .unwrap();
    let updates = events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 2);

    let remote_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_events);
    let _remote_subscription = target
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    assert!(
        target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("A2")).unwrap().input, "9");
    assert_eq!(remote_events.lock().unwrap().len(), 1);
    assert!(
        target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.model(), source.model());
    assert_eq!(remote_events.lock().unwrap().len(), 2);
    assert_eq!(
        remote_events.lock().unwrap()[0].origin,
        UpdateOrigin::Remote
    );
    assert!(
        !target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(remote_events.lock().unwrap().len(), 2);
}

#[test]
fn malformed_and_structural_remote_updates_roll_back_every_facade_state() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 701).unwrap();
    workbook.set_active_sheet(SheetId(1)).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A2"), "8", CalculationOptions::default())
        .unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "99".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    let assert_unchanged =
        |workbook: &Workbook,
         model: &WorkbookModel,
         state: &[u8],
         calculation: &betteroffice_xlsx::CalculationResult| {
            assert_eq!(workbook.model(), model);
            assert_eq!(workbook.encode_state_as_update_v1(), state);
            assert_eq!(workbook.active_sheet(), SheetId(1));
            assert!(workbook.can_undo());
            assert!(!workbook.can_redo());
            assert_eq!(workbook.proposals().len(), 1);
            assert_eq!(workbook.last_calculation(), calculation);
        };
    let model = workbook.model().clone();
    let state = workbook.encode_state_as_update_v1();
    let calculation = workbook.last_calculation().clone();
    assert!(matches!(
        workbook.apply_update_v1(&[0xff], CalculationOptions::default()),
        Err(Error::InvalidUpdate(_))
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);

    let mut structural = Workbook::open(&bytes).unwrap();
    structural
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let update = structural
        .encode_diff_v1(&workbook.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        workbook.apply_update_v1(&update, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);

    let mut shifted = Workbook::open(&bytes).unwrap();
    shifted
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let update = shifted
        .encode_diff_v1(&workbook.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        workbook.apply_update_v1(&update, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);
}

#[test]
fn rejected_update_preserves_unrelated_valid_causal_backlog() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open_collaborative(&bytes, 741).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 742).unwrap();
    let updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&updates);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    source
        .edit_cell(SheetId(0), cell("C3"), "one", CalculationOptions::default())
        .unwrap();
    source
        .edit_cell(SheetId(0), cell("C3"), "two", CalculationOptions::default())
        .unwrap();
    let updates = updates.lock().unwrap().clone();
    assert_eq!(updates.len(), 2);
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );

    let mut structural = Workbook::open(&bytes).unwrap();
    structural
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let invalid = structural
        .encode_diff_v1(&target.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        target.apply_update_v1(&invalid, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));

    assert!(
        target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C3")).unwrap().input, "two");
}

#[test]
fn independent_pending_chains_resolve_without_blocking_each_other() {
    let bytes = sample_xlsx();
    let mut first = Workbook::open_collaborative(&bytes, 743).unwrap();
    let mut second = Workbook::open_collaborative(&bytes, 744).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 745).unwrap();
    let first_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&first_updates);
    let _first_subscription = first
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let second_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&second_updates);
    let _second_subscription = second
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    first
        .edit_cell(SheetId(0), cell("C4"), "one", CalculationOptions::default())
        .unwrap();
    first
        .edit_cell(SheetId(0), cell("C4"), "two", CalculationOptions::default())
        .unwrap();
    second
        .edit_cell(
            SheetId(0),
            cell("C5"),
            "three",
            CalculationOptions::default(),
        )
        .unwrap();
    second
        .edit_cell(
            SheetId(0),
            cell("C5"),
            "four",
            CalculationOptions::default(),
        )
        .unwrap();
    let first_updates = first_updates.lock().unwrap().clone();
    let second_updates = second_updates.lock().unwrap().clone();

    assert!(
        !target
            .apply_update_v1(&first_updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&second_updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        target
            .apply_update_v1(&second_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C5")).unwrap().input, "four");
    assert_eq!(target.cell(SheetId(0), cell("C4")).unwrap().input, "");

    assert!(
        target
            .apply_update_v1(&first_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C4")).unwrap().input, "two");
}

#[test]
fn applicable_clients_in_a_partially_pending_update_are_not_blocked() {
    let bytes = sample_xlsx();
    let mut delayed = Workbook::open_collaborative(&bytes, 746).unwrap();
    let mut ready = Workbook::open_collaborative(&bytes, 747).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 748).unwrap();
    let delayed_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&delayed_updates);
    let _delayed_subscription = delayed
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let ready_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&ready_updates);
    let _ready_subscription = ready
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    delayed
        .edit_cell(SheetId(0), cell("D1"), "one", CalculationOptions::default())
        .unwrap();
    delayed
        .edit_cell(SheetId(0), cell("D1"), "two", CalculationOptions::default())
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D2"),
            "ready",
            CalculationOptions::default(),
        )
        .unwrap();
    let delayed_updates = delayed_updates.lock().unwrap().clone();
    let ready_updates = ready_updates.lock().unwrap().clone();
    let merged = YrsUpdate::merge_updates([
        YrsUpdate::decode_v1(&delayed_updates[1]).unwrap(),
        YrsUpdate::decode_v1(&ready_updates[0]).unwrap(),
    ])
    .encode_v1();

    assert!(
        target
            .apply_update_v1(&merged, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D2")).unwrap().input, "ready");
    assert_eq!(target.cell(SheetId(0), cell("D1")).unwrap().input, "");

    assert!(
        target
            .apply_update_v1(&delayed_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D1")).unwrap().input, "two");
}

#[test]
fn newly_applicable_clients_in_a_buffered_merged_update_are_committed() {
    let bytes = sample_xlsx();
    let mut delayed = Workbook::open_collaborative(&bytes, 749).unwrap();
    let mut ready = Workbook::open_collaborative(&bytes, 750).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 753).unwrap();
    let delayed_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&delayed_updates);
    let _delayed_subscription = delayed
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let ready_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&ready_updates);
    let _ready_subscription = ready
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    delayed
        .edit_cell(SheetId(0), cell("D3"), "one", CalculationOptions::default())
        .unwrap();
    delayed
        .edit_cell(SheetId(0), cell("D3"), "two", CalculationOptions::default())
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D4"),
            "three",
            CalculationOptions::default(),
        )
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D4"),
            "four",
            CalculationOptions::default(),
        )
        .unwrap();
    let delayed_updates = delayed_updates.lock().unwrap().clone();
    let ready_updates = ready_updates.lock().unwrap().clone();
    let merged = YrsUpdate::merge_updates([
        YrsUpdate::decode_v1(&delayed_updates[1]).unwrap(),
        YrsUpdate::decode_v1(&ready_updates[1]).unwrap(),
    ])
    .encode_v1();

    assert!(
        !target
            .apply_update_v1(&merged, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        target
            .apply_update_v1(&ready_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D4")).unwrap().input, "four");
    assert_eq!(target.cell(SheetId(0), cell("D3")).unwrap().input, "");
    let mut mirror = Workbook::open_collaborative(&bytes, 759).unwrap();
    mirror
        .apply_update_v1(
            &target.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(mirror.model(), target.model());

    assert!(
        target
            .apply_update_v1(&delayed_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D3")).unwrap().input, "two");
}

#[test]
fn wholly_pending_updates_do_not_reemit_existing_tombstones() {
    let bytes = sample_xlsx();
    let mut remote = Workbook::open_collaborative(&bytes, 754).unwrap();
    let mut local = Workbook::open_collaborative(&bytes, 755).unwrap();
    local
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "local",
            CalculationOptions::default(),
        )
        .unwrap();
    local
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A2"),
                    input: "proposal".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let remote_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_updates);
    let _remote_subscription = remote
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    remote
        .edit_cell(SheetId(0), cell("E1"), "one", CalculationOptions::default())
        .unwrap();
    remote
        .edit_cell(SheetId(0), cell("E1"), "two", CalculationOptions::default())
        .unwrap();
    let pending = remote_updates.lock().unwrap()[1].clone();
    let local_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&local_events);
    let _local_subscription = local
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    let state = local.encode_state_as_update_v1();

    let result = local
        .apply_update_v1(&pending, CalculationOptions::default())
        .unwrap();
    assert!(!result.applied);
    assert_eq!(local.encode_state_as_update_v1(), state);
    assert_eq!(local.proposals().len(), 1);
    assert!(local_events.lock().unwrap().is_empty());
}

#[test]
fn unresolved_invalid_updates_never_enter_live_yrs_state() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open(&bytes).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 751).unwrap();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    source
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Added".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    source
        .edit_cell(SheetId(1), cell("A1"), "17", CalculationOptions::default())
        .unwrap();
    let updates = events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 2);

    let state = target.encode_state_as_update_v1();
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.encode_state_as_update_v1(), state);
    assert!(matches!(
        target.apply_update_v1(&updates[0], CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_eq!(target.encode_state_as_update_v1(), state);
    assert_eq!(target.sheet_id("Data"), Some(SheetId(0)));

    let mut valid = Workbook::open_collaborative(&bytes, 752).unwrap();
    valid
        .edit_cell(SheetId(0), cell("A2"), "18", CalculationOptions::default())
        .unwrap();
    let update = valid
        .encode_diff_v1(&target.encode_state_vector_v1())
        .unwrap();
    assert!(
        target
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
}

#[test]
fn effective_remote_updates_clear_local_proposals() {
    let bytes = sample_xlsx();
    let mut remote = Workbook::open_collaborative(&bytes, 801).unwrap();
    let mut local = Workbook::open_collaborative(&bytes, 802).unwrap();
    local.set_active_sheet(SheetId(1)).unwrap();
    local
        .edit_cell(SheetId(0), cell("A2"), "11", CalculationOptions::default())
        .unwrap();
    assert!(local.can_undo());
    assert!(!local.can_redo());
    local
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "40".into(),
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    remote
        .edit_cell(SheetId(0), cell("A1"), "44", CalculationOptions::default())
        .unwrap();
    let update = remote
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert!(local.can_undo());
    assert!(!local.can_redo());
    assert!(local.proposals().is_empty());
    assert_eq!(local.active_sheet(), SheetId(1));
    assert!(local.undo(CalculationOptions::default()).unwrap().applied);
    assert_eq!(local.cell(SheetId(0), cell("A1")).unwrap().input, "44");
    assert_eq!(local.cell(SheetId(0), cell("A2")).unwrap().input, "5");
}

#[test]
fn update_observers_receive_one_owned_event_with_classified_origin() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 901).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 902).unwrap();
    let local_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&local_events);
    let local_subscription = left
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    left.edit_cells(
        SheetId(0),
        &[
            CellInput {
                cell: cell("A1"),
                input: "12".into(),
            },
            CellInput {
                cell: cell("A2"),
                input: "6".into(),
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    left.recalculate_all(CalculationOptions::default());
    let local_update = {
        let events = local_events.lock().unwrap();
        assert_eq!(events.len(), 1);
        assert_eq!(events[0].origin, UpdateOrigin::Local);
        events[0].update.clone()
    };

    let remote_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_events);
    let _remote_subscription = right
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    right
        .apply_update_v1(&local_update, CalculationOptions::default())
        .unwrap();
    let events = remote_events.lock().unwrap();
    assert_eq!(events.len(), 1);
    assert_eq!(events[0].origin, UpdateOrigin::Remote);
    assert_eq!(events[0].update, local_update);
    drop(events);

    drop(local_subscription);
    left.edit_cell(SheetId(0), cell("A1"), "13", CalculationOptions::default())
        .unwrap();
    assert_eq!(local_events.lock().unwrap().len(), 1);
}

#[test]
fn panicking_native_observers_do_not_split_authority_and_projection() {
    let bytes = sample_xlsx();
    let mut left =
        Workbook::open_collaborative_recalculated(&bytes, 911, CalculationOptions::default())
            .unwrap();
    let mut right =
        Workbook::open_collaborative_recalculated(&bytes, 912, CalculationOptions::default())
            .unwrap();
    let local_calls = Arc::new(AtomicUsize::new(0));
    let remote_calls = Arc::new(AtomicUsize::new(0));

    let _local_panic = left
        .observe_update_v1(|_| panic!("local observer panic"))
        .unwrap();
    let observed = Arc::clone(&local_calls);
    let _local_after = left
        .observe_update_v1(move |_| {
            observed.fetch_add(1, Ordering::SeqCst);
        })
        .unwrap();
    let _remote_panic = right
        .observe_update_v1(|_| panic!("remote observer panic"))
        .unwrap();
    let observed = Arc::clone(&remote_calls);
    let _remote_after = right
        .observe_update_v1(move |_| {
            observed.fetch_add(1, Ordering::SeqCst);
        })
        .unwrap();

    for (address, input) in [("C4", "first"), ("C5", "second")] {
        left.edit_cell(
            SheetId(0),
            cell(address),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
        let update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1()
        );
    }

    assert_eq!(local_calls.load(Ordering::SeqCst), 2);
    assert_eq!(remote_calls.load(Ordering::SeqCst), 2);
    assert_eq!(
        Workbook::open(&right.save().unwrap()).unwrap().model(),
        right.model()
    );
}

#[test]
fn collaborative_mode_rejects_all_structural_ops_before_mutation() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 1001).unwrap();
    let range = CellRange::new(cell("A1"), cell("A2"));
    let structural_ops = vec![
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::MergeCells {
            sheet: SheetId(0),
            range,
        },
        Op::UnmergeCells {
            sheet: SheetId(0),
            range,
        },
        Op::AddSheet {
            index: 1,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 1 },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".into(),
        },
        Op::RestoreSheet {
            sheet: SheetId(0),
            name: "Restored".into(),
            formulas: Vec::new(),
        },
    ];
    let model = workbook.model().clone();
    let state = workbook.encode_state_as_update_v1();
    for op in structural_ops {
        assert!(matches!(
            workbook.apply_ops(vec![op], CalculationOptions::default()),
            Err(Error::CollaborativeStructureOperation)
        ));
        assert_eq!(workbook.model(), &model);
        assert_eq!(workbook.encode_state_as_update_v1(), state);
        assert!(!workbook.can_undo());
    }

    assert!(
        workbook
            .apply_ops(
                vec![
                    Op::SetColWidth {
                        sheet: SheetId(0),
                        col: 0,
                        width: Some(22.0),
                    },
                    Op::SetRowHeight {
                        sheet: SheetId(0),
                        row: 0,
                        height: Some(24.0),
                    },
                ],
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
}

#[test]
fn collaboration_decoding_validates_malformed_and_oversized_payloads() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 1101).unwrap();
    assert!(matches!(
        workbook.encode_diff_v1(&[0xff]),
        Err(Error::InvalidStateVector(_))
    ));
    assert!(matches!(
        workbook.encode_diff_v1(&[0, 0]),
        Err(Error::InvalidStateVector(_))
    ));
    assert_eq!(MAX_COLLABORATION_STATE_VECTOR_ENTRIES, 65_536);
    assert!(matches!(
        workbook.encode_diff_v1(&[0x81, 0x80, 0x04]),
        Err(Error::InvalidStateVector(_))
    ));
    let oversized = vec![0_u8; MAX_COLLABORATION_BYTES + 1];
    assert!(matches!(
        workbook.encode_diff_v1(&oversized),
        Err(Error::CollaborationDataTooLarge { .. })
    ));
    assert!(matches!(
        workbook.apply_update_v1(&oversized, CalculationOptions::default()),
        Err(Error::CollaborationDataTooLarge { .. })
    ));
    assert!(matches!(
        Workbook::open_collaborative(&bytes, MAX_COLLABORATION_CLIENT_ID + 1),
        Err(Error::InvalidClientId { .. })
    ));
    let max_client = Workbook::open_collaborative(&bytes, MAX_COLLABORATION_CLIENT_ID).unwrap();
    assert_eq!(max_client.client_id(), MAX_COLLABORATION_CLIENT_ID);
}
