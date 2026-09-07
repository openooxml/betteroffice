use betteroffice_xlsx::{
    Cell, CellRange, CellRef, CellValue, DrawCmd, Error, FreezePane, GridGeometry, PrintMetrics,
    Sheet, SheetId, Viewport, Workbook, WorkbookModel,
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

fn workbook() -> Workbook {
    let mut sheet = Sheet::new("Data");
    sheet.col_widths.insert(0, 12.0);
    sheet.col_widths.insert(1, 13.0);
    sheet.row_heights.insert(1, 24.0);
    sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(1, 1)));
    for row in 0..4 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Text {
                    value: format!("Row {row}"),
                },
                ..Cell::default()
            },
        );
    }
    Workbook::from_model(WorkbookModel {
        sheets: vec![sheet, Sheet::new("Other")],
        ..WorkbookModel::default()
    })
    .unwrap()
}

#[test]
fn printing_uses_font_device_metrics_without_changing_screen_or_source() {
    let mut workbook = workbook();
    workbook.set_active_sheet(SheetId(1)).unwrap();
    let model = workbook.model().clone();
    let saved = workbook.save().unwrap();
    let screen_geometry = GridGeometry::new(&model.sheets[0]);
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 300.0,
        height: 200.0,
    };
    let screen = workbook.display_list_for(SheetId(0), &viewport).unwrap();
    let printed = workbook
        .print_display_list(
            SheetId(0),
            CellRange::parse_a1("A2:B3").unwrap(),
            &metrics(),
            true,
        )
        .unwrap();
    assert!((printed.width - 604.0 / 3.0).abs() < 0.001);
    assert!((printed.height - 52.0).abs() < 0.001);
    assert_eq!(printed.grid.start_row, 1);
    assert_eq!(printed.grid.col_offsets[1], 96.0);
    assert!(printed.commands.iter().any(|c| matches!(c, DrawCmd::Text { text, font_family, .. } if text == "Row 1" && font_family.as_deref() == Some("Calibri"))));
    assert!(
        !printed
            .commands
            .iter()
            .any(|c| matches!(c, DrawCmd::Text { text, .. } if text == "Row 0" || text == "Row 3"))
    );
    assert!(printed.commands.iter().any(|c| matches!(c, DrawCmd::Line { color, width, .. } if color == "#000000" && (*width - 4.0 / 3.0).abs() < 0.001)));
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.active_sheet(), SheetId(1));
    assert_eq!(workbook.save().unwrap(), saved);
    assert_eq!(
        workbook.display_list_for(SheetId(0), &viewport).unwrap(),
        screen
    );
    assert_eq!(
        GridGeometry::new(&model.sheets[0]).col_x(1),
        screen_geometry.col_x(1)
    );
}

#[test]
fn stored_column_width_includes_padding_and_device_rounding() {
    let mut metrics = metrics();
    metrics.dpi = 96.0;
    metrics.max_digit_width = 7.0;
    assert_eq!(metrics.column_pixels(8.7109375), 61.0);
    assert_eq!(metrics.column_pixels(0.0), 0.0);
    metrics.dpi = 72.0;
    metrics.max_digit_width = 6.0;
    assert_eq!(metrics.column_pixels(12.0), 96.0);
}

#[test]
fn print_gridlines_are_optional_and_invalid_metrics_are_rejected() {
    let workbook = workbook();
    let range = CellRange::parse_a1("A1:B3").unwrap();
    let printed = workbook
        .print_display_list(SheetId(0), range, &metrics(), false)
        .unwrap();
    assert!(
        !printed
            .commands
            .iter()
            .any(|c| matches!(c, DrawCmd::Line { .. }))
    );
    for dpi in [0.0, -1.0, f32::NAN, f32::INFINITY, 601.0] {
        let mut invalid = metrics();
        invalid.dpi = dpi;
        assert!(matches!(
            workbook.print_display_list(SheetId(0), range, &invalid, true),
            Err(Error::InvalidViewport)
        ));
    }
    assert!(matches!(
        workbook.print_display_list(
            SheetId(0),
            CellRange::parse_a1("A1:XFD1048576").unwrap(),
            &metrics(),
            true
        ),
        Err(Error::DisplayTooLarge { .. })
    ));
}
