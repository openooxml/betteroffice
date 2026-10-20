use betteroffice_xlsx::{
    Cell, CellRange, CellRef, CellValue, DisplayList, DrawCmd, Error, FreezePane, GridGeometry,
    MAX_COLS, MAX_DISPLAY_CELLS, MAX_ROWS, PrintMetrics, Sheet, SheetId, Viewport, Workbook,
    WorkbookModel,
};

const SCREEN: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 1600.0,
    height: 1000.0,
};

fn open(sheet: Sheet) -> Workbook {
    Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap()
}

fn text(sheet: &mut Sheet, row: u32, col: u32) {
    sheet.set_cell(
        CellRef::new(row, col),
        Cell {
            value: CellValue::Text {
                value: format!("r{row}c{col}"),
            },
            ..Cell::default()
        },
    );
}

fn filtered_rows() -> Sheet {
    let mut sheet = Sheet::new("Filtered");
    for row in 1..=100_002 {
        if row != 50_000 && row != 100_000 {
            sheet.row_heights.insert(row, 0.0);
        }
    }
    for row in [0, 50_000, 100_000] {
        for col in 0..13 {
            text(&mut sheet, row, col);
        }
    }
    text(&mut sheet, 1, 0);
    sheet
}

fn hidden_columns() -> Sheet {
    let mut sheet = Sheet::new("Hidden columns");
    for col in 1..MAX_COLS {
        if col != 8_000 && col != 16_000 {
            sheet.col_widths.insert(col, 0.0);
        }
    }
    for row in 0..20 {
        for col in [0, 8_000, 16_000] {
            text(&mut sheet, row, col);
        }
    }
    text(&mut sheet, 0, 1);
    sheet
}

fn has_text(frame: &DisplayList, row: u32, col: u32) -> bool {
    let wanted = format!("r{row}c{col}");
    frame
        .commands
        .iter()
        .any(|command| matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == wanted))
}

fn assert_too_large(result: betteroffice_xlsx::Result<DisplayList>) {
    assert!(matches!(
        result,
        Err(Error::DisplayTooLarge { cells, max }) if cells > max && max == MAX_DISPLAY_CELLS
    ));
}

#[test]
fn display_cap_excludes_hidden_rows_and_frozen_rows() {
    for frozen in [0, 50_001] {
        let mut sheet = filtered_rows();
        sheet.freeze_pane = Some(FreezePane::new(frozen, 0, CellRef::new(frozen, 0)));
        let frame = open(sheet).display_list(&SCREEN).unwrap();
        for row in [0, 50_000, 100_000] {
            for col in 0..13 {
                assert!(has_text(&frame, row, col));
            }
        }
        assert!(!has_text(&frame, 1, 0));
        assert!(frame.grid.row_offsets.len() < 60);
        let indices = frame.grid.row_indices.as_ref().unwrap();
        assert_eq!(&indices[..3], &[0, 50_000, 100_000]);
    }
}

#[test]
fn display_cap_excludes_hidden_columns_and_frozen_columns() {
    for frozen in [0, 8_001] {
        let mut sheet = hidden_columns();
        sheet.freeze_pane = Some(FreezePane::new(0, frozen, CellRef::new(0, frozen)));
        let frame = open(sheet).display_list(&SCREEN).unwrap();
        for row in 0..20 {
            for col in [0, 8_000, 16_000] {
                assert!(has_text(&frame, row, col));
            }
        }
        assert!(!has_text(&frame, 0, 1));
        assert_eq!(frame.grid.col_offsets.len(), 4);
        assert_eq!(
            frame.grid.col_indices.as_deref(),
            Some(&[0, 8_000, 16_000][..])
        );
    }
}

#[test]
fn display_cap_skips_a_million_implicit_hidden_rows() {
    let mut sheet = Sheet::new("Sparse visible rows");
    sheet.format.zero_height = true;
    for row in [0, 500_000, MAX_ROWS - 1] {
        sheet.row_heights.insert(row, 15.0);
        text(&mut sheet, row, 0);
    }
    let frame = open(sheet).display_list(&SCREEN).unwrap();
    assert_eq!(frame.grid.row_offsets.len(), 4);
    assert_eq!(
        frame.grid.row_indices.as_deref(),
        Some(&[0, 500_000, MAX_ROWS - 1][..])
    );
    for row in [0, 500_000, MAX_ROWS - 1] {
        assert!(has_text(&frame, row, 0));
    }
}

#[test]
fn oversized_frames_degrade_at_the_scroll_origin_with_frozen_panes() {
    for frozen in [false, true] {
        let mut sheet = Sheet::new("Dense");
        sheet.format.default_row_height_pt = Some(3.0);
        sheet.format.custom_height = true;
        for col in 0..MAX_COLS {
            sheet.col_widths.insert(col, 0.5);
        }
        if frozen {
            sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(1, 1)));
        }
        let geometry = GridGeometry::new(&sheet, &WorkbookModel::default().styles);
        for (row, col) in [(0, 0), (101, 101)] {
            text(&mut sheet, row, col);
        }
        let workbook = open(sheet);
        for (row, col) in [(0, 0), (100, 100)] {
            let viewport = Viewport {
                x: geometry.col_x(col),
                y: geometry.row_y(row),
                width: SCREEN.width / 0.1,
                height: SCREEN.height / 0.1,
            };
            let frame = workbook.display_list(&viewport).unwrap();
            assert!(frame.truncated);
            assert_eq!(
                (frame.width, frame.height),
                (viewport.width, viewport.height)
            );
            let cells = (frame.grid.row_offsets.len() - 1) * (frame.grid.col_offsets.len() - 1);
            assert!(cells > 0 && cells as u64 <= MAX_DISPLAY_CELLS);
            assert!(has_text(
                &frame,
                if row == 0 { 0 } else { 101 },
                if col == 0 { 0 } else { 101 }
            ));
            if frozen {
                assert!(has_text(&frame, 0, 0));
                assert_eq!(
                    frame.grid.row_indices.as_ref().map(|rows| rows[1]),
                    (row > 0).then_some(101)
                );
            } else {
                assert_eq!(frame.grid.start_row, row);
                assert_eq!(frame.grid.start_col, col);
            }
            assert!(frame.grid.row_offsets.last().unwrap() < &viewport.height);
            assert!(
                workbook
                    .visible_merged_ranges(SheetId(0), &viewport)
                    .is_ok()
            );
        }
    }
}

#[test]
fn display_cap_degrades_excessive_visible_rows_and_columns() {
    for (sheet, height) in [
        (filtered_rows(), 6_000_000.0),
        (hidden_columns(), 2_000_000.0),
    ] {
        let frame = open(sheet)
            .display_list(&Viewport { height, ..SCREEN })
            .unwrap();
        assert!(frame.truncated);
        assert!(
            (frame.grid.row_offsets.len() - 1) as u64 * (frame.grid.col_offsets.len() - 1) as u64
                <= MAX_DISPLAY_CELLS
        );
    }
}

#[cfg(feature = "raster")]
#[test]
fn raster_calls_still_reject_excessive_visible_cells() {
    let mut sheet = Sheet::new("Dense");
    sheet.format.default_row_height_pt = Some(0.5);
    sheet.format.custom_height = true;
    for col in 0..MAX_COLS {
        sheet.col_widths.insert(col, 0.25);
    }
    let workbook = open(sheet);
    assert!(matches!(
        workbook.render_png(&SCREEN),
        Err(Error::DisplayTooLarge { .. })
    ));
    assert!(matches!(
        workbook.render_sheet(
            SheetId(0),
            &betteroffice_xlsx::RenderOptions {
                range: Some(CellRange::new(CellRef::new(0, 0), CellRef::new(999, 999))),
                scale: 0.1,
                ..Default::default()
            }
        ),
        Err(Error::DisplayTooLarge { .. })
    ));
}

#[test]
fn capped_frames_clip_merges() {
    let mut sheet = Sheet::new("Objects");
    let crossing = CellRange::new(CellRef::new(0, 0), CellRef::new(100, 0));
    sheet.merges.push(crossing);
    sheet
        .merges
        .push(CellRange::new(CellRef::new(100, 1), CellRef::new(101, 1)));
    text(&mut sheet, 0, 0);
    let workbook = open(sheet);
    let viewport = Viewport {
        width: 1_000_000.0,
        ..SCREEN
    };
    let frame = workbook.display_list(&viewport).unwrap();
    assert!(frame.truncated);
    let bottom = *frame.grid.row_offsets.last().unwrap();
    assert_eq!(
        workbook
            .visible_merged_ranges(SheetId(0), &viewport)
            .unwrap(),
        vec![crossing]
    );
    assert!(frame.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text { text, clip, .. }
            if text.as_ref() == "r0c0" && clip.y + clip.h <= bottom
    )));
}

#[test]
fn hidden_merge_anchors_still_draw_text() {
    for (hide_row, hide_col) in [(true, false), (false, true), (true, true)] {
        let mut sheet = Sheet::new("Merged");
        sheet
            .merges
            .push(CellRange::new(CellRef::new(0, 0), CellRef::new(2, 2)));
        if hide_row {
            sheet.row_heights.insert(0, 0.0);
        }
        if hide_col {
            sheet.col_widths.insert(0, 0.0);
        }
        text(&mut sheet, 0, 0);
        let workbook = open(sheet);
        assert!(has_text(&workbook.display_list(&SCREEN).unwrap(), 0, 0));
        let printed = workbook
            .print_display_list(
                SheetId(0),
                CellRange::new(CellRef::new(0, 0), CellRef::new(2, 2)),
                &metrics(),
                false,
            )
            .unwrap();
        assert!(has_text(&printed, 0, 0));
    }
}

#[test]
fn hidden_columns_do_not_interrupt_text_spill() {
    let mut sheet = Sheet::new("Spill");
    sheet.col_widths.insert(1, 0.0);
    text(&mut sheet, 0, 0);
    text(&mut sheet, 0, 4);
    let workbook = open(sheet);
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let frame = workbook.display_list(&SCREEN).unwrap();
    let clip = frame.commands.iter().find_map(|command| match command {
        DrawCmd::Text { text, clip, .. } if text.as_ref() == "r0c0" => Some(clip),
        _ => None,
    });
    let clip = clip.unwrap();
    assert_eq!(clip.x + clip.w, geometry.col_x(4));
}

fn metrics() -> PrintMetrics {
    PrintMetrics {
        dpi: 96.0,
        max_digit_width: 7.0,
        default_row_height_pt: 15.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    }
}

#[test]
fn print_cap_excludes_hidden_rows() {
    let workbook = open(filtered_rows());
    let frame = workbook
        .print_display_list(
            SheetId(0),
            CellRange::new(CellRef::new(0, 0), CellRef::new(100_002, 12)),
            &metrics(),
            true,
        )
        .unwrap();
    for row in [0, 50_000, 100_000] {
        for col in 0..13 {
            assert!(has_text(&frame, row, col));
        }
    }
    assert!(!has_text(&frame, 1, 0));
    assert!(frame.grid.row_offsets.len() <= 5);
}

#[test]
fn print_cap_excludes_hidden_columns() {
    let workbook = open(hidden_columns());
    let frame = workbook
        .print_display_list(
            SheetId(0),
            CellRange::new(CellRef::new(0, 0), CellRef::new(19, MAX_COLS - 1)),
            &metrics(),
            true,
        )
        .unwrap();
    for row in 0..20 {
        for col in [0, 8_000, 16_000] {
            assert!(has_text(&frame, row, col));
        }
    }
    assert!(!has_text(&frame, 0, 1));
    assert_eq!(frame.grid.col_offsets.len(), 4);
}

#[test]
fn print_cap_still_rejects_excessive_visible_cells() {
    for sheet in [filtered_rows(), hidden_columns()] {
        assert_too_large(open(sheet).print_display_list(
            SheetId(0),
            CellRange::new(CellRef::new(0, 0), CellRef::new(MAX_ROWS - 1, MAX_COLS - 1)),
            &metrics(),
            false,
        ));
    }
}
