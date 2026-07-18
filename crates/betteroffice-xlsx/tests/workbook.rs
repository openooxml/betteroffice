#[cfg(feature = "raster")]
use betteroffice_xlsx::RenderOptions;
use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellValue, DrawCmd, Error, Op,
    ProposalEditInput, ProposalRequest, Sheet, SheetId, Viewport, Workbook, WorkbookModel,
};

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
