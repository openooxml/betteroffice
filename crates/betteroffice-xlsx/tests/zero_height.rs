use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRef, DrawCmd, GridGeometry, SheetId, Viewport, Workbook,
    WorkbookModel,
};

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(betteroffice_xlsx::Sheet::new("Data"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="12.75" zeroHeight="1"/><sheetData><row r="1" ht="15" customHeight="1"/><row r="2"><c r="A2" t="inlineStr"><is><t>visible</t></is></c></row><row r="4" hidden="1"/></sheetData></worksheet>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

const SCREEN: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 1600.0,
    height: 1000.0,
};

#[test]
fn zero_height_frame_draws_unsized_row_text() {
    let workbook = Workbook::open(&source()).unwrap();
    let frame = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    assert!(frame.commands.iter().any(|command| {
        matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "visible")
    }));
}

#[test]
fn zero_height_short_sheet_does_not_exceed_display_cell_cap() {
    let workbook = Workbook::open(&source()).unwrap();
    let frame = workbook.display_list_for(SheetId(0), &SCREEN).unwrap();
    assert_eq!(frame.grid.row_offsets, [0.0, 20.0, 37.0]);
}

#[test]
fn zero_height_snapshots_and_peer_edits_keep_content_rows_visible() {
    use betteroffice_xlsx::{SnapshotBudget, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder};

    let source = source();
    let mut workbook = Workbook::open(&source).unwrap();
    let mut peer = Workbook::open_collaborative(&source, 811).unwrap();
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(1, 0),
                input: "edited".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    peer.apply_update_v1(
        &workbook.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    assert!(!peer.sheet(SheetId(0)).unwrap().row_hidden(1));
    assert!(peer.sheet(SheetId(0)).unwrap().row_hidden(2));

    let initial = Workbook::open(&source).unwrap();
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&initial, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&initial).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(restored.model(), initial.model());
    assert_eq!(
        restored
            .display_list_for(SheetId(0), &SCREEN)
            .unwrap()
            .grid
            .row_offsets,
        [0.0, 20.0, 37.0],
    );
}

#[test]
fn zero_height_empty_row_stays_hidden() {
    let workbook = Workbook::open(&source()).unwrap();
    let sheet = workbook.sheet(SheetId(0)).unwrap();
    assert!(sheet.cell(CellRef::new(2, 0)).is_none());
    assert!(!sheet.row_heights.contains_key(&2));
    assert!(!sheet.row_styles.contains_key(&2));
    assert!(sheet.row_hidden(2));
}

#[test]
fn zero_height_cached_geometry_tracks_cell_edits_undo_and_redo() {
    let mut workbook = Workbook::open(&source()).unwrap();
    let offsets = |workbook: &Workbook| {
        workbook
            .display_list_for(SheetId(0), &SCREEN)
            .unwrap()
            .grid
            .row_offsets
    };
    let original = offsets(&workbook);
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(2, 0),
                input: "added".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let edited = offsets(&workbook);
    assert_eq!(edited, [0.0, 20.0, 37.0, 54.0]);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(offsets(&workbook), original);
    assert!(workbook.sheet(SheetId(0)).unwrap().row_hidden(2));
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(offsets(&workbook), edited);
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(2, 0),
                input: String::new(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(offsets(&workbook), original);
    assert!(workbook.sheet(SheetId(0)).unwrap().row_hidden(2));
}

#[test]
fn zero_height_edited_save_keeps_unsized_row_visible_without_size_attributes() {
    let source = source();
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.save().unwrap(), source);
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: CellRef::new(1, 0),
                input: "edited".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(xml.contains(r#"<row r="2">"#));
    let reopened = Workbook::open(&saved).unwrap();
    let sheet = reopened.sheet(SheetId(0)).unwrap();
    assert!(!sheet.row_hidden(1));
    assert!(sheet.row_hidden(2));
    assert!(sheet.row_hidden(3));
    let geometry = GridGeometry::new(sheet, &reopened.model().styles);
    assert_eq!(geometry.row_y(2) - geometry.row_y(1), 17.0);
    let frame = reopened.display_list_for(SheetId(0), &SCREEN).unwrap();
    assert!(frame.commands.iter().any(|command| {
        matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "edited")
    }));
}
