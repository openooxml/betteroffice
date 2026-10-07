use betteroffice_xlsx::{
    Cell, CellRef, CellValue, DrawCmd, Error, GridGeometry, MAX_COLS, MAX_DISPLAY_CELLS, MAX_ROWS,
    Sheet, SheetId, Viewport, Workbook, WorkbookModel,
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
    let geometry = GridGeometry::new(workbook.sheet(SheetId(0)).unwrap(), &workbook.model().styles);
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
    let geometry = GridGeometry::new(workbook.sheet(SheetId(0)).unwrap(), &workbook.model().styles);
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
