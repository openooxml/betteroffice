use betteroffice_xlsx::{
    Cell, CellRange, CellRef, CellValue, DrawCmd, Error, FreezePane, GridGeometry, Hyperlink,
    PrintMetrics, Sheet, SheetId, Viewport, Workbook, WorkbookModel,
};
use xlsx_model::styles::{Border, BorderEdge, BorderStyle, Color, Fill, Stylesheet, Xf};

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

#[test]
fn partial_merged_print_ranges_clip_without_moving_the_anchor() {
    let mut sheet = Sheet::new("Merged");
    sheet.merges.push(CellRange::parse_a1("A1:C3").unwrap());
    for col in 0..3 {
        sheet.col_widths.insert(col, 12.0);
    }
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Text {
                value: "Merged content".into(),
            },
            style: Some(0),
            ..Cell::default()
        },
    );
    let edge = BorderEdge {
        style: BorderStyle::Thin,
        color: Some(Color::Rgb("#0000ff".into())),
    };
    let mut styles = Stylesheet::default();
    styles.fills = vec![Fill::Solid(Color::Rgb("#ffd700".into()))];
    styles.borders = vec![Border {
        left: Some(edge.clone()),
        right: Some(edge.clone()),
        top: Some(edge.clone()),
        bottom: Some(edge),
    }];
    styles.cell_xfs = vec![Xf {
        fill: Some(0),
        border: Some(0),
        ..Xf::default()
    }];
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        styles,
        ..WorkbookModel::default()
    })
    .unwrap();
    let render = |range| {
        workbook
            .print_display_list(
                SheetId(0),
                CellRange::parse_a1(range).unwrap(),
                &metrics(),
                false,
            )
            .unwrap()
    };
    let full = render("A1:C3");
    let partial = render("B2:C3");
    let texts = |list: &betteroffice_xlsx::DisplayList| {
        list.commands
            .iter()
            .filter_map(|cmd| {
                if let DrawCmd::Text { x, y, clip, .. } = cmd {
                    Some((*x, *y, *clip))
                } else {
                    None
                }
            })
            .collect::<Vec<_>>()
    };
    let full_text = texts(&full);
    let partial_text = texts(&partial);
    assert_eq!(full_text.len(), 1);
    assert_eq!(partial_text.len(), 1);
    assert!((partial_text[0].0 - (full_text[0].0 - 96.0)).abs() < 0.001);
    assert!((partial_text[0].1 - (full_text[0].1 - 56.0 / 3.0)).abs() < 0.001);
    assert_eq!(partial_text[0].2.x, 0.0);
    assert_eq!(partial_text[0].2.y, 0.0);
    assert_eq!(partial_text[0].2.w, partial.width);
    assert_eq!(partial_text[0].2.h, partial.height);
    assert!(partial.commands.iter().any(|cmd| matches!(cmd,
        DrawCmd::FillRect { x, y, w, h, color, .. }
        if *x == 0.0 && *y == 0.0 && *w == partial.width && *h == partial.height && color == "#ffd700"
    )));
    assert_eq!(
        partial
            .commands
            .iter()
            .filter(|cmd| matches!(cmd,
                DrawCmd::Line { color, .. } if color == "#0000ff"
            ))
            .count(),
        2
    );
}

#[test]
fn display_only_hyperlinks_use_the_same_print_font_and_position_as_cells() {
    let mut sheet = Sheet::new("Links");
    sheet.set_cell(CellRef::new(0, 0), Cell::default());
    for col in 0..2 {
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::new(CellRef::new(0, col), CellRef::new(0, col)),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: Some("Link".into()),
        });
    }
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    let mut metrics = metrics();
    metrics.font_family = "Example Font".into();
    metrics.font_size_pt = 16.0;
    metrics.default_row_height_pt = 24.0;
    metrics.font_ascent = 15.0;
    metrics.font_descent = 4.0;
    let printed = workbook
        .print_display_list(
            SheetId(0),
            CellRange::parse_a1("A1:B1").unwrap(),
            &metrics,
            false,
        )
        .unwrap();
    let texts: Vec<_> = printed
        .commands
        .iter()
        .filter_map(|cmd| {
            if let DrawCmd::Text {
                x,
                y,
                font_family,
                font_size,
                underline,
                ..
            } = cmd
            {
                assert_eq!(font_family.as_deref(), Some("Example Font"));
                assert_eq!(*font_size, 16.0);
                assert!(*underline);
                Some((*x, *y))
            } else {
                None
            }
        })
        .collect();
    assert_eq!(texts.len(), 2);
    assert_eq!(texts[0].1, texts[1].1);
    assert!((texts[1].0 - texts[0].0 - printed.grid.col_offsets[1]).abs() < 0.001);
    assert_eq!(texts[0].0, 4.0);
    assert!((texts[0].1 - 80.0 / 3.0).abs() < 0.001);
}
