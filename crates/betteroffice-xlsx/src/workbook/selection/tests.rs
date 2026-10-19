use super::*;
use xlsx_model::Cell;

fn workbook() -> Workbook {
    let mut sheet = Sheet::new("Data");
    for row in 0..6 {
        for col in 0..4 {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number { value: 0.0 },
                    ..Cell::default()
                },
            );
        }
    }
    for (row, col) in [(2, 7), (2, 8), (2, 9), (9, 2), (10, 2), (11, 2), (179, 39)] {
        sheet.set_cell(
            CellRef::new(row, col),
            Cell {
                value: CellValue::Bool { value: false },
                ..Cell::default()
            },
        );
    }
    Workbook::from_model(WorkbookModel {
        sheets: vec![sheet, Sheet::new("Empty")],
        ..WorkbookModel::default()
    })
    .unwrap()
}

fn edge(workbook: &Workbook, row: u32, col: u32, direction: SelectionDirection) -> Option<CellRef> {
    workbook
        .selection_data(SheetId(0), CellRef::new(row, col), Some(direction))
        .unwrap()
        .edge
}

#[test]
fn contiguous_runs_gaps_and_empty_starts_in_all_directions() {
    use SelectionDirection::*;
    let workbook = workbook();
    for (row, col, direction, expected) in [
        (2, 2, Right, Some((2, 3))),
        (2, 3, Right, Some((2, 7))),
        (2, 7, Right, Some((2, 9))),
        (2, 6, Right, Some((2, 7))),
        (2, 5, Right, Some((2, 7))),
        (2, 2, Left, Some((2, 0))),
        (2, 7, Left, Some((2, 3))),
        (2, 6, Left, Some((2, 3))),
        (2, 5, Left, Some((2, 3))),
        (2, 2, Down, Some((5, 2))),
        (5, 2, Down, Some((9, 2))),
        (8, 2, Down, Some((9, 2))),
        (7, 2, Down, Some((9, 2))),
        (9, 2, Down, Some((11, 2))),
        (4, 2, Up, Some((0, 2))),
        (9, 2, Up, Some((5, 2))),
        (8, 2, Up, Some((5, 2))),
        (13, 2, Up, Some((11, 2))),
        (2, 0, Left, None),
        (0, 2, Up, None),
        (2, 9, Right, None),
        (11, 2, Down, None),
        (0, 39, Down, Some((179, 39))),
        (179, 0, Right, Some((179, 39))),
    ] {
        assert_eq!(
            edge(&workbook, row, col, direction),
            expected.map(|(r, c)| CellRef::new(r, c))
        );
    }
}

#[test]
fn formulas_and_values_count_but_formatting_does_not() {
    use SelectionDirection::*;
    let mut sheet = Sheet::new("Data");
    for col in [0, 1, 3] {
        sheet.set_cell(
            CellRef::new(0, col),
            Cell {
                formula: Some("IF(TRUE,\"\",1)".to_owned()),
                ..Cell::default()
            },
        );
    }
    sheet.set_cell(
        CellRef::new(MAX_ROWS - 1, MAX_COLS - 1),
        Cell {
            value: CellValue::Text {
                value: String::new(),
            },
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::new(CellRef::new(0, 2), CellRef::new(0, 2)),
            StylePatch {
                font_size: Some(18.0),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(edge(&workbook, 0, 0, Right), Some(CellRef::new(0, 1)));
    assert_eq!(edge(&workbook, 0, 1, Right), Some(CellRef::new(0, 3)));
    assert_eq!(edge(&workbook, 0, 2, Right), Some(CellRef::new(0, 3)));
    assert_eq!(
        edge(&workbook, 0, MAX_COLS - 1, Down),
        Some(CellRef::new(MAX_ROWS - 1, MAX_COLS - 1))
    );
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(0, 0), None)
            .unwrap()
            .region,
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 0,
            right: 1
        })
    );
}

#[test]
fn current_region_includes_adjacent_cells_and_stops_at_empty_separators() {
    let workbook = workbook();
    let region = Some(SelectionRegion {
        top: 0,
        left: 0,
        bottom: 5,
        right: 3,
    });
    for at in [CellRef::new(1, 1), CellRef::new(5, 3)] {
        assert_eq!(
            workbook
                .selection_data(SheetId(0), at, None)
                .unwrap()
                .region,
            region
        );
    }
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(6, 4), None)
            .unwrap()
            .region,
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 6,
            right: 4
        })
    );
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(20, 20), None)
            .unwrap()
            .region,
        None
    );
    assert_eq!(
        workbook
            .selection_data(SheetId(1), CellRef::new(0, 0), None)
            .unwrap()
            .region,
        None
    );
}

#[test]
fn region_grows_the_bounding_box_iteratively_including_diagonals() {
    let mut sheet = Sheet::new("Data");
    for (row, col) in [(0, 5), (1, 4), (2, 2), (3, 3), (4, 4)] {
        sheet.set_cell(
            CellRef::new(row, col),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(2, 2), None)
            .unwrap()
            .region,
        Some(SelectionRegion {
            top: 0,
            left: 2,
            bottom: 4,
            right: 5
        })
    );
}

#[test]
fn queries_are_read_only_and_validate_addresses() {
    let workbook = workbook();
    let before = workbook.save().unwrap();
    let version = workbook.version();
    workbook
        .selection_data(SheetId(1), CellRef::new(0, 0), None)
        .unwrap();
    assert!(
        workbook
            .selection_data(SheetId(2), CellRef::new(0, 0), None)
            .is_err()
    );
    assert!(
        workbook
            .selection_data(SheetId(0), CellRef::new(MAX_ROWS, 0), None)
            .is_err()
    );
    assert!(
        workbook
            .selection_data(SheetId(0), CellRef::new(0, MAX_COLS), None)
            .is_err()
    );
    assert_eq!(workbook.active_sheet(), SheetId(0));
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), before);
}

#[test]
fn region_follows_a_large_reverse_sparse_staircase() {
    let mut sheet = Sheet::new("Data");
    for row in 0..100_000 {
        sheet.set_cell(
            CellRef::new(row, row % 2),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    sheet.set_cell(
        CellRef::new(MAX_ROWS - 1, MAX_COLS - 1),
        Cell {
            value: CellValue::Bool { value: true },
            ..Cell::default()
        },
    );
    assert_eq!(
        current_region(&sheet, CellRef::new(99_999, 1)),
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 99_999,
            right: 1,
        })
    );
}

#[test]
fn merged_cells_are_filled_and_directional_queries_start_at_their_edges() {
    use SelectionDirection::*;
    let mut sheet = Sheet::new("Data");
    sheet
        .merges
        .push(CellRange::new(CellRef::new(2, 2), CellRef::new(4, 4)));
    for (row, col) in [
        (0, 3),
        (1, 3),
        (5, 3),
        (6, 3),
        (3, 0),
        (3, 1),
        (3, 5),
        (3, 6),
    ] {
        sheet.set_cell(
            CellRef::new(row, col),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    for (direction, expected) in [
        (Up, (0, 3)),
        (Down, (6, 3)),
        (Left, (3, 0)),
        (Right, (3, 6)),
    ] {
        assert_eq!(
            edge(&workbook, 3, 3, direction),
            Some(CellRef::new(expected.0, expected.1))
        );
    }
    assert_eq!(edge(&workbook, 2, 0, Right), Some(CellRef::new(2, 2)));
    assert_eq!(edge(&workbook, 0, 2, Down), Some(CellRef::new(2, 2)));
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(3, 3), None)
            .unwrap()
            .region,
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 6,
            right: 6,
        })
    );
}

#[test]
fn a_merge_bridges_regions_without_expanding_its_cells() {
    let mut sheet = Sheet::new("Data");
    sheet.merges.push(CellRange::new(
        CellRef::new(0, 1),
        CellRef::new(MAX_ROWS - 1, MAX_COLS - 2),
    ));
    for at in [CellRef::new(0, 0), CellRef::new(MAX_ROWS - 1, MAX_COLS - 1)] {
        sheet.set_cell(
            at,
            Cell {
                value: CellValue::Bool { value: true },
                ..Cell::default()
            },
        );
    }
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    assert_eq!(
        edge(&workbook, 0, 2, SelectionDirection::Left),
        Some(CellRef::new(0, 0))
    );
    assert_eq!(
        workbook
            .selection_data(SheetId(0), CellRef::new(0, 0), None)
            .unwrap()
            .region,
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: MAX_ROWS - 1,
            right: MAX_COLS - 1,
        })
    );
}

#[test]
fn hidden_tracks_do_not_end_runs_or_supply_data_targets() {
    use SelectionDirection::*;
    let mut sheet = Sheet::new("Data");
    sheet.row_heights.insert(1, 0.0);
    sheet.col_widths.insert(1, 0.0);
    for (row, col) in [(0, 0), (0, 2), (0, 3), (2, 0), (3, 0), (0, 5), (5, 0)] {
        sheet.set_cell(
            CellRef::new(row, col),
            Cell {
                value: CellValue::Bool { value: true },
                ..Cell::default()
            },
        );
    }
    sheet.col_widths.insert(5, 0.0);
    sheet.row_heights.insert(5, 0.0);
    for (at, direction, expected) in [
        ((0, 0), Right, Some((0, 3))),
        ((0, 3), Left, Some((0, 0))),
        ((0, 0), Down, Some((3, 0))),
        ((3, 0), Up, Some((0, 0))),
        ((0, 3), Right, None),
        ((3, 0), Down, None),
        ((0, 4), Right, None),
        ((4, 0), Down, None),
    ] {
        assert_eq!(
            data_edge(&sheet, CellRef::new(at.0, at.1), direction),
            expected.map(|(row, col)| CellRef::new(row, col))
        );
    }
    assert_eq!(
        current_region(&sheet, CellRef::new(0, 0)),
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 3,
            right: 3,
        })
    );
}

#[test]
fn hidden_cells_do_not_bridge_visible_empty_separators() {
    let mut sheet = Sheet::new("Data");
    for (row, col) in [(0, 0), (0, 2), (0, 3), (2, 0), (3, 0)] {
        sheet.set_cell(
            CellRef::new(row, col),
            Cell {
                value: CellValue::Bool { value: true },
                ..Cell::default()
            },
        );
    }
    sheet.col_widths.insert(2, 0.0);
    sheet.row_heights.insert(2, 0.0);
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 0), SelectionDirection::Right),
        Some(CellRef::new(0, 3))
    );
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 0), SelectionDirection::Down),
        Some(CellRef::new(3, 0))
    );
    assert_eq!(
        current_region(&sheet, CellRef::new(0, 0)),
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 0,
            right: 0,
        })
    );
}

#[test]
fn hidden_merge_boundaries_are_trimmed_without_expanding_cells() {
    let mut sheet = Sheet::new("Data");
    sheet
        .merges
        .push(CellRange::new(CellRef::new(0, 1), CellRef::new(3, 4)));
    for index in [1, 4] {
        sheet.col_widths.insert(index, 0.0);
    }
    sheet.row_heights.insert(3, 0.0);
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 0), SelectionDirection::Right),
        Some(CellRef::new(0, 2))
    );
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 5), SelectionDirection::Left),
        Some(CellRef::new(0, 3))
    );
    assert_eq!(
        current_region(&sheet, CellRef::new(0, 2)),
        Some(SelectionRegion {
            top: 0,
            left: 2,
            bottom: 2,
            right: 3,
        })
    );
    sheet.col_widths.insert(2, 0.0);
    sheet.col_widths.insert(3, 0.0);
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 0), SelectionDirection::Right),
        None
    );
    assert_eq!(current_region(&sheet, CellRef::new(0, 0)), None);
}

#[test]
fn zero_height_defaults_skip_large_hidden_spans() {
    let mut sheet = Sheet::new("Data");
    sheet.format.zero_height = true;
    for row in [0, 900_000] {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                formula: Some("1".to_owned()),
                ..Cell::default()
            },
        );
    }
    assert_eq!(
        data_edge(&sheet, CellRef::new(0, 0), SelectionDirection::Down),
        Some(CellRef::new(900_000, 0))
    );
    assert_eq!(
        data_edge(&sheet, CellRef::new(900_000, 0), SelectionDirection::Up),
        Some(CellRef::new(0, 0))
    );
    assert_eq!(
        current_region(&sheet, CellRef::new(0, 0)),
        Some(SelectionRegion {
            top: 0,
            left: 0,
            bottom: 900_000,
            right: 0,
        })
    );
}
