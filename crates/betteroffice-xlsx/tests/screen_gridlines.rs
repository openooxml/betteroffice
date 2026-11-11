use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRange, CellRef, DisplayList, DrawCmd, PrintMetrics, Sheet,
    SheetId, SnapshotBudget, Viewport, Workbook, WorkbookModel, WorkbookSnapshotBuilder,
    WorkbookSnapshotEncoder,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Dashboard"));
    model.sheets.push(Sheet::new("Data"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    for (path, views) in [
        (
            "xl/worksheets/sheet1.xml",
            r#"<sheetView showGridLines="0" workbookViewId="0"/><sheetView workbookViewId="1"/>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            r#"<sheetView workbookViewId="0"/><sheetView showGridLines="0" workbookViewId="1"/>"#,
        ),
    ] {
        parts.iter_mut().find(|(part, _)| part == path).unwrap().1 = format!(
            r#"<worksheet xmlns="{MAIN}"><sheetViews>{views}</sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>text</t></is></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

const SCREEN: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 400.0,
    height: 200.0,
};

fn metrics() -> PrintMetrics {
    PrintMetrics {
        dpi: 72.0,
        max_digit_width: 6.0,
        default_row_height_pt: 14.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    }
}

fn lines(frame: &DisplayList) -> usize {
    frame
        .commands
        .iter()
        .filter(|command| matches!(command, DrawCmd::Line { .. }))
        .count()
}

fn screen_lines(workbook: &Workbook, sheet: u32) -> usize {
    lines(&workbook.display_list_for(SheetId(sheet), &SCREEN).unwrap())
}

#[test]
fn the_first_sheet_view_hides_screen_gridlines_without_changing_print() {
    let workbook = Workbook::open(&source()).unwrap();
    let dashboard = workbook.sheet(SheetId(0)).unwrap().format;
    assert!(dashboard.hide_gridlines);
    assert_eq!(dashboard.default_row_height_pt, Some(15.0));
    assert!(!workbook.sheet(SheetId(1)).unwrap().format.hide_gridlines);
    assert_eq!(screen_lines(&workbook, 0), 0);
    assert!(screen_lines(&workbook, 1) > 0);

    let range = CellRange::parse_a1("A1:C3").unwrap();
    for sheet in [SheetId(0), SheetId(1)] {
        let print = |gridlines| {
            lines(
                &workbook
                    .print_display_list(sheet, range, &metrics(), gridlines)
                    .unwrap(),
            )
        };
        assert!(print(true) > 0);
        assert_eq!(print(false), 0);
    }
}

#[test]
fn hidden_gridlines_survive_save_snapshots_peers_and_recovery() {
    let source = source();
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.save().unwrap(), source);

    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(restored.model(), workbook.model());
    assert_eq!(screen_lines(&restored, 0), 0);

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
    assert_eq!(screen_lines(&peer, 0), 0);
    assert!(screen_lines(&peer, 1) > 0);

    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    let (recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(recovered.model(), workbook.model());
    assert_eq!(screen_lines(&recovered, 0), 0);

    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert!(reopened.sheet(SheetId(0)).unwrap().format.hide_gridlines);
    assert!(!reopened.sheet(SheetId(1)).unwrap().format.hide_gridlines);
}
