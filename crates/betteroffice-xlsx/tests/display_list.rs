use betteroffice_xlsx::{
    Cell, CellRef, CellValue, DrawCmd, Error, FreezePane, GridGeometry, MAX_COLS,
    MAX_DISPLAY_CELLS, MAX_ROWS, Sheet, SheetId, Viewport, Workbook, WorkbookModel,
};

fn workbook() -> Workbook {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        CellRef::new(MAX_ROWS - 1, MAX_COLS - 1),
        Cell {
            value: CellValue::Text {
                value: "edge".into(),
            },
            ..Cell::default()
        },
    );
    Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap()
}

#[test]
fn display_lists_clamp_past_grid_end() {
    let workbook = workbook();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let x = geometry.col_x(MAX_COLS - 1);
    let y = geometry.row_y(MAX_ROWS - 1);
    let col_width = geometry.col_x(MAX_COLS) - x;
    let row_height = geometry.row_y(MAX_ROWS) - y;
    for (width, height) in [
        (col_width * 4.0, row_height / 2.0),
        (col_width / 2.0, row_height * 4.0),
        (f32::MAX / 4.0, f32::MAX / 4.0),
    ] {
        let frame = workbook
            .display_list(&Viewport {
                x,
                y,
                width,
                height,
            })
            .unwrap();
        assert_eq!(frame.grid.start_row, MAX_ROWS - 1);
        assert_eq!(frame.grid.start_col, MAX_COLS - 1);
        assert_eq!(frame.grid.row_offsets, vec![0.0, row_height]);
        assert_eq!(frame.grid.col_offsets, vec![0.0, col_width]);
        assert!(frame.commands.iter().any(|command| {
            matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "edge")
        }));
    }
}

#[test]
fn display_lists_beyond_grid_end_have_no_tracks() {
    let workbook = workbook();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let right = geometry.col_x(MAX_COLS) + 128.0;
    let bottom = geometry.row_y(MAX_ROWS) + 40.0;
    for (x, y) in [(right, bottom), (right, 0.0), (0.0, bottom)] {
        let frame = workbook
            .display_list(&Viewport {
                x,
                y,
                width: 200.0,
                height: 100.0,
            })
            .unwrap();
        if x == right {
            assert!(frame.grid.col_offsets.is_empty());
        }
        if y == bottom {
            assert!(frame.grid.row_offsets.is_empty());
        }
        assert_eq!(frame.commands.len(), 1);
        assert!(matches!(frame.commands[0], DrawCmd::FillRect { .. }));
    }
}

#[test]
fn display_lists_custom_dimensions_clamp_at_grid_boundaries() {
    let mut sheet = Sheet::new("Data");
    sheet.col_widths.insert(3, 8.44);
    sheet.row_heights.insert(3, 15.1);
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let right = geometry.col_x(MAX_COLS);
    let bottom = geometry.row_y(MAX_ROWS);
    for (x, y, width, height) in [
        (right, 0.0, 200.0, 6_000_000.0),
        (right + 128.0, 0.0, 200.0, 6_000_000.0),
        (0.0, bottom, right, 100.0),
        (0.0, bottom + 40.0, right, 100.0),
    ] {
        let viewport = Viewport {
            x,
            y,
            width,
            height,
        };
        let (rows, cols) = geometry.viewport_range(&viewport);
        let frame = workbook.display_list(&viewport).unwrap();
        if x >= right {
            assert_eq!(cols, MAX_COLS..MAX_COLS);
            assert!(frame.grid.col_offsets.is_empty());
        }
        if y >= bottom {
            assert_eq!(rows, MAX_ROWS..MAX_ROWS);
            assert!(frame.grid.row_offsets.is_empty());
        }
        assert_eq!(frame.commands.len(), 1);
        assert!(matches!(frame.commands[0], DrawCmd::FillRect { .. }));
    }

    let viewport = Viewport {
        x: right - 32.0,
        y: bottom - 8.0,
        width: 32.0,
        height: 8.0,
    };
    assert_eq!(
        geometry.viewport_range(&viewport),
        (MAX_ROWS - 1..MAX_ROWS, MAX_COLS - 1..MAX_COLS)
    );
    let frame = workbook.display_list(&viewport).unwrap();
    assert_eq!(frame.grid.start_row, MAX_ROWS - 1);
    assert_eq!(frame.grid.start_col, MAX_COLS - 1);
    assert_eq!(frame.grid.row_offsets.len(), 2);
    assert_eq!(frame.grid.col_offsets.len(), 2);
    assert_eq!(frame.grid.row_offsets[1], viewport.height);
    assert_eq!(frame.grid.col_offsets[1], viewport.width);
}

#[test]
fn display_lists_preserve_last_row_strip_with_frozen_pane() {
    let mut sheet = Sheet::new("Data");
    sheet.row_heights.insert(0, 14.25);
    sheet.freeze_pane = Some(FreezePane::new(1, 0, CellRef::new(MAX_ROWS - 1, 0)));
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    assert_eq!(geometry.row_y(1), 19.0);
    assert_eq!(geometry.row_y(MAX_ROWS), 20_971_520.0);
    let viewport = Viewport {
        x: 0.0,
        y: 20_971_500.0,
        width: 100.0,
        height: 20.0,
    };
    assert_eq!(geometry.viewport_range(&viewport).0, MAX_ROWS - 1..MAX_ROWS);
    let frame = workbook.display_list(&viewport).unwrap();
    assert_eq!(frame.grid.start_row, 0);
    assert_eq!(
        frame.grid.row_indices.as_deref(),
        Some(&[0, MAX_ROWS - 1][..])
    );
    assert_eq!(frame.grid.row_offsets, vec![0.0, 19.0, 20.0]);
    assert!(frame.commands.iter().any(|command| {
        matches!(
            command,
            DrawCmd::Line { x1, y1, x2, y2, .. }
                if x1 != x2 && *y1 == 20.0 && *y2 == 20.0
        )
    }));
}

#[test]
fn display_lists_count_cells_when_a_start_rounds_to_the_grid_end() {
    let mut sheet = Sheet::new("Data");
    sheet.col_widths.insert(0, 8.35);
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let viewport = Viewport {
        x: geometry.col_x(MAX_COLS) - 0.125,
        y: 0.0,
        width: 0.125,
        height: 6_000_000.0,
    };
    assert_eq!(geometry.viewport_range(&viewport).1, MAX_COLS - 1..MAX_COLS);
    assert!(matches!(
        workbook.display_list(&viewport),
        Err(Error::DisplayTooLarge { cells, max }) if cells > max && max == MAX_DISPLAY_CELLS
    ));
}

#[test]
fn display_lists_past_one_grid_end_build_no_tracks_on_the_other_axis() {
    let workbook = workbook();
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let right = geometry.col_x(MAX_COLS);
    let bottom = geometry.row_y(MAX_ROWS);
    for viewport in [
        Viewport {
            x: right,
            y: 0.0,
            width: 200.0,
            height: bottom,
        },
        Viewport {
            x: 0.0,
            y: bottom,
            width: right,
            height: 100.0,
        },
    ] {
        let frame = workbook.display_list(&viewport).unwrap();
        assert!(frame.grid.row_offsets.is_empty());
        assert!(frame.grid.col_offsets.is_empty());
        assert_eq!(frame.commands.len(), 1);
    }
}

#[test]
fn display_lists_reject_excessive_in_sheet_spans() {
    let workbook = workbook();
    let error = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 6_000_000.0,
        })
        .unwrap_err();
    assert!(matches!(
        error,
        Error::DisplayTooLarge { cells, max } if cells > max && max == MAX_DISPLAY_CELLS
    ));
}
