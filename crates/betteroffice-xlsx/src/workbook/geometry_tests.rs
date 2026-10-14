use super::*;
use xlsx_model::Cell;
use xlsx_model::styles::{Font, Xf};
use xlsx_render::geometry::{geometry_counters, geometry_prefix_counters, reset_geometry_counters};

fn model(rows: u32, cols: u32) -> WorkbookModel {
    let mut sheet = Sheet::new("Data");
    sheet.format.default_row_height_pt = Some(15.0);
    for row in 0..rows {
        for col in 0..cols {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
        }
    }
    WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    }
}

fn assert_full_equivalence(workbook: &Workbook) {
    for index in 0..workbook.model.sheets.len() {
        let id = SheetId(index as u32);
        let actual = workbook.sheet_geometry(id).unwrap();
        let full = GridGeometry::new(workbook.sheet(id).unwrap(), &workbook.model.styles);
        for row in (0..80).chain([1_000, MAX_ROWS]) {
            assert_eq!(actual.row_y(row), full.row_y(row), "row {row}");
            if row < MAX_ROWS {
                assert_eq!(
                    actual.row_y(row + 1) - actual.row_y(row),
                    full.row_y(row + 1) - full.row_y(row)
                );
                if actual.row_y(row + 1) > actual.row_y(row) {
                    assert_eq!(actual.row_at_y(actual.row_y(row)), row, "row {row}");
                    assert_eq!(full.row_at_y(full.row_y(row)), row, "row {row}");
                }
            }
            for y in [
                actual.row_y(row) - 0.01,
                actual.row_y(row),
                actual.row_y(row) + 0.01,
            ] {
                assert_eq!(actual.row_at_y(y), full.row_at_y(y));
            }
        }
        for col in (0..15).chain([MAX_COLS]) {
            assert_eq!(actual.col_x(col), full.col_x(col));
            if col < MAX_COLS && actual.col_x(col + 1) > actual.col_x(col) {
                assert_eq!(actual.col_at_x(actual.col_x(col)), col, "col {col}");
                assert_eq!(full.col_at_x(full.col_x(col)), col, "col {col}");
            }
            for x in [
                actual.col_x(col) - 0.01,
                actual.col_x(col),
                actual.col_x(col) + 0.01,
            ] {
                assert_eq!(actual.col_at_x(x), full.col_at_x(x));
            }
        }
        for y in [0.0_f32, 100.0, 500.0, 1_000.0] {
            let bottom = actual.row_y(MAX_ROWS);
            if bottom == 0.0 {
                continue;
            }
            let y = y.min((bottom - 1.0).max(0.0));
            let viewport = Viewport {
                x: 10.0,
                y,
                width: 400.0,
                height: 250.0_f32.min(bottom - y),
            };
            assert_eq!(
                actual.viewport_range(&viewport),
                full.viewport_range(&viewport)
            );
            assert_eq!(
                serde_json::to_value(workbook.display_list_for(id, &viewport).unwrap()).unwrap(),
                serde_json::to_value(
                    xlsx_render::build_display_list(workbook.model(), id, &viewport).unwrap()
                )
                .unwrap()
            );
        }
    }
}

#[test]
fn seeded_mutations_match_full_geometry() {
    let mut input = model(48, 6);
    input.sheets[0].format.default_row_height_pt = Some(15.5);
    input.styles.fonts = [11.0, 18.0, 32.0]
        .into_iter()
        .map(|size| Font {
            size_pt: Some(size),
            ..Font::default()
        })
        .collect();
    input.styles.cell_xfs = (0..3)
        .map(|font| Xf {
            font: Some(font),
            ..Xf::default()
        })
        .collect();
    let mut other = Sheet::new("Other");
    other.format.zero_height = true;
    other.row_heights.insert(2, 20.0);
    input.sheets.push(other);
    let mut fractional = Sheet::new("Fractional");
    fractional.format.custom_height = true;
    fractional.format.default_row_height_pt = Some(15.25);
    fractional.row_heights.insert(0, 25.25);
    input.sheets.push(fractional);
    let mut workbook = Workbook::from_model(input).unwrap();
    let options = CalculationOptions::default();
    assert_full_equivalence(&workbook);
    let mut seed = 0x6d2b_79f5_u64;
    for step in 0..160 {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        let row = (seed % 48) as u32;
        let col = ((seed >> 8) % 6) as u32;
        let sheet = SheetId(0);
        let at = CellRef::new(row, col);
        let range = CellRange::new(at, CellRef::new(row, (col + 1).min(5)));
        let op = match step % 16 {
            0..=3 => Op::SetCell {
                sheet,
                at,
                cell: CellState {
                    value: CellValue::Number { value: step as f64 },
                    style: Some((step % 3) as u32),
                    ..CellState::default()
                },
            },
            4 => Op::SetCell {
                sheet,
                at,
                cell: CellState::default(),
            },
            5 => Op::PatchRangeStyle {
                sheet,
                range,
                patch: StylePatch {
                    font_size: Some([9.0, 12.0, 40.0][(seed % 3) as usize]),
                    ..StylePatch::default()
                },
            },
            6 => Op::PatchRangeStyle {
                sheet,
                range,
                patch: StylePatch {
                    text_wrapping: Some(TextWrapping::Wrap),
                    ..StylePatch::default()
                },
            },
            7 => Op::SetRowHeight {
                sheet,
                row,
                height: Some(if seed & 1 == 0 { 0.0 } else { 25.25 }),
            },
            8 => Op::SetRowHeight {
                sheet,
                row,
                height: None,
            },
            9 => Op::SetColWidth {
                sheet,
                col,
                width: Some(if seed & 1 == 0 { 0.0 } else { 13.25 }),
            },
            10 => Op::MergeCells {
                sheet,
                range: CellRange::new(CellRef::new(row, 0), CellRef::new(row + 1, 2)),
            },
            11 => Op::UnmergeCells {
                sheet,
                range: CellRange::new(CellRef::new(row, 0), CellRef::new(row + 1, 2)),
            },
            12 => Op::InsertRows {
                sheet,
                at: row,
                count: 1,
            },
            13 => Op::DeleteRows {
                sheet,
                at: row,
                count: 1,
            },
            14 => Op::InsertCols {
                sheet,
                at: col,
                count: 1,
            },
            _ => Op::DeleteCols {
                sheet,
                at: col,
                count: 1,
            },
        };
        reset_geometry_counters();
        let result = workbook.apply_ops(vec![op], options).unwrap();
        workbook.cell_rect(sheet, at).unwrap();
        assert_eq!(
            geometry_counters().0,
            u64::from(result.applied && step % 16 >= 12)
        );
        assert_full_equivalence(&workbook);
        if step % 9 == 0 {
            workbook
                .edit_cells(
                    sheet,
                    &[
                        CellInput {
                            cell: CellRef::new(row, 0),
                            input: "17".into(),
                        },
                        CellInput {
                            cell: CellRef::new(row, 1),
                            input: "".into(),
                        },
                    ],
                    options,
                )
                .unwrap();
            assert_full_equivalence(&workbook);
            workbook.undo(options).unwrap();
            assert_full_equivalence(&workbook);
            workbook.redo(options).unwrap();
            assert_full_equivalence(&workbook);
        }
    }
}

#[test]
fn enter_cell_rect_visits_only_its_row() {
    for rows in [100_000, 200_000] {
        let mut workbook = Workbook::from_model(model(rows, 11)).unwrap();
        let sheet = SheetId(0);
        let at = CellRef::new(rows - 2, 5);
        workbook.cell_rect(sheet, at).unwrap();
        let retained = workbook.sheet_geometry(sheet).unwrap();
        let old_offset = retained.row_y(rows);
        for input in ["2", "3", "", "4"] {
            reset_geometry_counters();
            workbook
                .edit_cell(sheet, at, input, CalculationOptions::default())
                .unwrap();
            workbook.cell_rect(sheet, at).unwrap();
            let (full, visits) = geometry_counters();
            assert_eq!(full, 0);
            assert!(visits <= 11, "visited {visits} cells");
            let (updates, queries) = geometry_prefix_counters();
            assert!(updates <= 21, "updated {updates} nodes");
            assert!(queries <= 84, "queried {queries} nodes");
            reset_geometry_counters();
            for row in [0, rows / 2, rows - 1, MAX_ROWS - 1] {
                workbook.cell_rect(sheet, CellRef::new(row, 10)).unwrap();
            }
            assert_eq!(geometry_counters(), (0, 0));
        }
        assert_eq!(retained.row_y(rows), old_offset);
    }
}

#[test]
fn changing_row_height_updates_only_logarithmic_prefixes() {
    for rows in [100_000, 200_000] {
        let mut workbook = Workbook::from_model(model(rows, 1)).unwrap();
        let sheet = SheetId(0);
        let at = CellRef::new(rows / 2, 0);
        workbook.cell_rect(sheet, at).unwrap();
        let retained = workbook.sheet_geometry(sheet).unwrap();
        let before = retained.row_y(at.row);
        let after = retained.row_y(at.row + 1);
        let bottom = retained.row_y(rows);
        for input in ["", "1", "", "1"] {
            reset_geometry_counters();
            workbook
                .edit_cell(sheet, at, input, CalculationOptions::default())
                .unwrap();
            workbook.cell_rect(sheet, at).unwrap();
            let (full, visits) = geometry_counters();
            assert_eq!(full, 0);
            assert_eq!(visits, u64::from(!input.is_empty()));
            let (updates, queries) = geometry_prefix_counters();
            assert!(updates > 0 && updates <= 21, "updated {updates} nodes");
            assert!(queries <= 84, "queried {queries} nodes");
            let actual = workbook.sheet_geometry(sheet).unwrap();
            let change = if input.is_empty() {
                xlsx_render::geometry::row_pt_to_px(15.0)
                    - xlsx_render::geometry::row_pt_to_px(14.0)
            } else {
                0.0
            };
            assert_eq!(actual.row_y(at.row), before);
            assert!((actual.row_y(at.row + 1) - after - change).abs() < 0.25);
            assert!((actual.row_y(rows) - bottom - change).abs() < 0.25);
            if input.is_empty() {
                assert!(actual.row_y(at.row + 1) > after);
                assert!(actual.row_y(rows) > bottom);
            }
            assert_eq!(retained.row_y(rows), bottom);
        }
    }
}

#[test]
fn structural_changes_rebuild_once() {
    let mut workbook = Workbook::from_model(model(30, 11)).unwrap();
    let sheet = SheetId(0);
    let at = CellRef::new(2, 2);
    workbook.cell_rect(sheet, at).unwrap();
    for op in [
        Op::InsertRows {
            sheet,
            at: 2,
            count: 1,
        },
        Op::DeleteRows {
            sheet,
            at: 2,
            count: 1,
        },
        Op::InsertCols {
            sheet,
            at: 2,
            count: 1,
        },
        Op::DeleteCols {
            sheet,
            at: 2,
            count: 1,
        },
        Op::AddSheet {
            index: 1,
            name: "Other".into(),
        },
        Op::RemoveSheet { index: 1 },
    ] {
        reset_geometry_counters();
        workbook
            .apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
        workbook.cell_rect(sheet, at).unwrap();
        assert_eq!(geometry_counters().0, 1);
        workbook.cell_rect(sheet, at).unwrap();
        assert_eq!(geometry_counters().0, 1);
        assert_full_equivalence(&workbook);
    }
}

#[test]
fn replacement_changes_rebuild_once() {
    let mut workbook = Workbook::from_model(model(20, 4)).unwrap();
    let sheet = SheetId(0);
    let at = CellRef::new(2, 2);
    workbook.cell_rect(sheet, at).unwrap();
    for kind in 0..4 {
        workbook.sheet_info().unwrap();
        let mut replacement = workbook.model().clone();
        match kind {
            0 => replacement.sheets[0].format.default_row_height_pt = Some(18.0),
            1 => replacement.sheets[0].format.zero_height = true,
            2 => replacement.styles.fonts.push(Font {
                size_pt: Some(22.0),
                ..Font::default()
            }),
            _ => replacement.sheets[0].format.custom_height = true,
        }
        reset_geometry_counters();
        if kind == 2 {
            workbook.install_model_with_ops(replacement, &[]).unwrap();
        } else {
            workbook.install_model(replacement).unwrap();
        }
        workbook.cell_rect(sheet, at).unwrap();
        assert_eq!(geometry_counters().0, 1);
        workbook.cell_rect(sheet, at).unwrap();
        assert_eq!(geometry_counters().0, 1);
        assert_full_equivalence(&workbook);
        let cold = Workbook::from_model(workbook.model().clone()).unwrap();
        assert_eq!(workbook.sheet_info().unwrap(), cold.sheet_info().unwrap());
    }
}

#[test]
fn tallest_contributors_merges_and_snapshots_stay_consistent() {
    let mut input = model(10, 4);
    input.styles.fonts = [11.0, 32.0]
        .into_iter()
        .map(|size| Font {
            size_pt: Some(size),
            ..Font::default()
        })
        .collect();
    input.styles.cell_xfs = (0..2)
        .map(|font| Xf {
            font: Some(font),
            ..Xf::default()
        })
        .collect();
    input.sheets[0].col_styles.push(xlsx_model::ColStyle {
        first: 3,
        last: 3,
        xf: 1,
    });
    input.sheets[0].col_widths.insert(3, 0.0);
    let mut workbook = Workbook::from_model(input).unwrap();
    let options = CalculationOptions::default();
    let sheet = SheetId(0);
    let retained = workbook.sheet_geometry(sheet).unwrap();
    let old_offset = retained.row_y(10);
    for ops in [
        vec![
            Op::SetCell {
                sheet,
                at: CellRef::new(2, 0),
                cell: CellState {
                    style: Some(1),
                    ..CellState::default()
                },
            },
            Op::SetCell {
                sheet,
                at: CellRef::new(2, 1),
                cell: CellState {
                    style: Some(1),
                    ..CellState::default()
                },
            },
        ],
        vec![Op::SetCell {
            sheet,
            at: CellRef::new(2, 0),
            cell: CellState::default(),
        }],
        vec![Op::SetCell {
            sheet,
            at: CellRef::new(2, 1),
            cell: CellState {
                style: Some(0),
                ..CellState::default()
            },
        }],
        vec![Op::SetCell {
            sheet,
            at: CellRef::new(2, 3),
            cell: CellState::default(),
        }],
        vec![Op::MergeCells {
            sheet,
            range: CellRange::parse_a1("D4:D8").unwrap(),
        }],
        vec![Op::MergeCells {
            sheet,
            range: CellRange::parse_a1("B7:D9").unwrap(),
        }],
        vec![Op::UnmergeCells {
            sheet,
            range: CellRange::parse_a1("B7:D9").unwrap(),
        }],
    ] {
        reset_geometry_counters();
        workbook.apply_ops(ops, options).unwrap();
        workbook.cell_rect(sheet, CellRef::new(2, 0)).unwrap();
        assert_eq!(geometry_counters().0, 0);
        assert_full_equivalence(&workbook);
        assert_eq!(retained.row_y(10), old_offset);
    }
    reset_geometry_counters();
    let update = workbook
        .commit_user(&[Op::RestoreColStyles {
            sheet,
            styles: Vec::new(),
        }])
        .unwrap();
    workbook.rebuild_and_recalculate(options);
    workbook.publish(update);
    workbook.cell_rect(sheet, CellRef::new(2, 0)).unwrap();
    assert_eq!(geometry_counters().0, 0);
    assert_full_equivalence(&workbook);
}

#[test]
fn pending_edits_and_rejections_keep_other_sheets_warm() {
    let mut input = model(20, 4);
    input.sheets[0].format.default_row_height_pt = None;
    let mut other = input.sheets[0].clone();
    other.name = "Other".into();
    input.sheets.push(other);
    let mut workbook = Workbook::from_model(input).unwrap();
    let sheet = SheetId(0);
    let at = CellRef::new(2, 0);
    workbook.sheet_geometry(sheet).unwrap();
    let other = workbook.sheet_geometry(SheetId(1)).unwrap();
    for value in ["2", "3", "4"] {
        workbook
            .edit_cell(sheet, at, value, CalculationOptions::default())
            .unwrap();
    }
    assert!(
        workbook
            .apply_ops(
                vec![Op::SetRowHeight {
                    sheet,
                    row: 2,
                    height: Some(-1.0)
                }],
                CalculationOptions::default()
            )
            .is_err()
    );
    reset_geometry_counters();
    assert!(Arc::ptr_eq(
        &other,
        &workbook.sheet_geometry(SheetId(1)).unwrap()
    ));
    assert_eq!(geometry_counters(), (0, 0));
    workbook.cell_rect(sheet, at).unwrap();
    assert_eq!(geometry_counters(), (0, 4));
    assert_full_equivalence(&workbook);
    reset_geometry_counters();
    workbook
        .apply_ops(
            vec![Op::SetFreezePane {
                sheet,
                pane: Some(FreezePane::new(1, 1, CellRef::new(1, 1))),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.cell_rect(sheet, at).unwrap();
    assert_eq!(geometry_counters(), (0, 0));
}

#[test]
fn cross_sheet_spill_changes_update_geometry() {
    let mut input = model(1, 1);
    input.styles.fonts = [11.0, 32.0]
        .into_iter()
        .map(|size| Font {
            size_pt: Some(size),
            ..Font::default()
        })
        .collect();
    input.styles.cell_xfs = (0..2)
        .map(|font| Xf {
            font: Some(font),
            ..Xf::default()
        })
        .collect();
    let mut other = Sheet::new("Other");
    other.format.default_row_height_pt = Some(15.0);
    other.col_styles.push(xlsx_model::ColStyle {
        first: 0,
        last: 0,
        xf: 1,
    });
    let anchor = CellRef::new(3, 0);
    other.set_cell(
        anchor,
        Cell {
            formula: Some("_xlfn.SEQUENCE(Data!A1)".into()),
            ..Cell::default()
        },
    );
    other.set_array_formula(anchor, CellRange::new(anchor, anchor));
    input.sheets.push(other);
    let mut workbook = Workbook::from_model(input).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    assert_full_equivalence(&workbook);
    for value in ["4", "2", "6", "3"] {
        reset_geometry_counters();
        let result = workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                value,
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(result.changed.iter().any(|cell| cell.sheet == SheetId(1)));
        workbook.cell_rect(SheetId(1), CellRef::new(8, 0)).unwrap();
        assert_eq!(geometry_counters().0, 0);
        assert_eq!(
            workbook
                .sheet(SheetId(1))
                .unwrap()
                .array_formula(anchor)
                .unwrap()
                .end
                .row,
            3 + value.parse::<u32>().unwrap() - 1
        );
        assert_full_equivalence(&workbook);
    }
}

fn warm_edit_reads(workbook: &Workbook, at: CellRef) -> (SheetInfo, xlsx_render::Rect) {
    let info = workbook.sheet_info().unwrap();
    assert_eq!(workbook.sheet_info().unwrap(), info);
    let rect = workbook.cell_rect(SheetId(0), at).unwrap();
    let viewport = Viewport {
        x: rect.x,
        y: rect.y,
        width: 640.0,
        height: 400.0,
    };
    let frame = workbook.display_list_for(SheetId(0), &viewport).unwrap();
    let cached = workbook.display_list_for(SheetId(0), &viewport).unwrap();
    assert_eq!(
        serde_json::to_vec(&frame).unwrap(),
        serde_json::to_vec(&cached).unwrap()
    );
    (info, rect)
}

fn warm_edit_counters() -> (u64, u64, u64) {
    (
        xlsx_render::geometry::geometry_row_storage_allocations(),
        xlsx_model::workbook::used_range_scans(),
        geometry_counters().0,
    )
}

fn reset_warm_edit_counters() {
    reset_geometry_counters();
    xlsx_model::workbook::reset_used_range_scans();
}

fn assert_warm_edit_bounds(workbook: &Workbook, expected: Option<CellRange>) {
    assert_eq!(
        workbook
            .sheet_info_cache
            .lock()
            .unwrap()
            .as_ref()
            .unwrap()
            .bounds,
        expected
    );
    assert_eq!(workbook.sheet(SheetId(0)).unwrap().used_range(), expected);
    let geometry = workbook.sheet_geometry(SheetId(0)).unwrap();
    let content = sheet_content(expected, None, &geometry);
    let info = workbook.sheet_info().unwrap();
    assert_eq!(info.content_width, content.width);
    assert_eq!(info.content_height, content.height);
}

fn assert_cold_geometry(workbook: &Workbook, at: CellRef) {
    let cold = Workbook::from_model(workbook.model().clone()).unwrap();
    let actual = workbook.sheet_info().unwrap();
    let expected = cold.sheet_info().unwrap();
    assert_eq!(actual.content_height, expected.content_height);
    assert_eq!(actual.content_width, expected.content_width);
    assert_eq!(
        workbook.cell_rect(SheetId(0), at).unwrap(),
        cold.cell_rect(SheetId(0), at).unwrap()
    );
}

#[test]
fn warm_edit_numeric_keeps_storage_and_bounds() {
    let mut workbook = Workbook::from_model(model(100_000, 11)).unwrap();
    let at = CellRef::new(50_000, 5);
    let before = warm_edit_reads(&workbook, at);
    for input in ["2", "3"] {
        reset_warm_edit_counters();
        workbook
            .edit_cell(SheetId(0), at, input, CalculationOptions::default())
            .unwrap();
        let after = warm_edit_reads(&workbook, at);
        assert_eq!(warm_edit_counters(), (0, 0, 0));
        assert_eq!(before, after);
        assert_warm_edit_bounds(
            &workbook,
            Some(CellRange::new(CellRef::new(0, 0), CellRef::new(99_999, 10))),
        );
    }
}

#[test]
fn warm_edit_height_keeps_storage_and_bounds() {
    let mut input = model(100_000, 11);
    let at = CellRef::new(50_000, 5);
    for col in (0..11).filter(|col| *col != at.col) {
        input.sheets[0].set_cell(CellRef::new(at.row, col), Cell::default());
    }
    let mut workbook = Workbook::from_model(input).unwrap();
    let before = warm_edit_reads(&workbook, at);
    for input in ["", "1"] {
        reset_warm_edit_counters();
        workbook
            .edit_cell(SheetId(0), at, input, CalculationOptions::default())
            .unwrap();
        let after = warm_edit_reads(&workbook, at);
        assert_eq!(warm_edit_counters(), (0, 0, 0));
        if input.is_empty() {
            assert!(after.1.h > before.1.h);
            assert!(after.0.content_height > before.0.content_height);
        } else {
            assert_eq!(after, before);
        }
        assert_warm_edit_bounds(
            &workbook,
            Some(CellRange::new(CellRef::new(0, 0), CellRef::new(99_999, 10))),
        );
    }
}

#[test]
fn warm_edit_dependent_keeps_storage_and_bounds() {
    let mut input = model(100_000, 11);
    input.sheets[0]
        .cell_mut(CellRef::new(0, 0))
        .unwrap()
        .formula = Some("B1".into());
    let mut workbook = Workbook::from_model(input).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    let at = CellRef::new(0, 1);
    let before = warm_edit_reads(&workbook, at);
    for input in ["2", "3"] {
        reset_warm_edit_counters();
        let result = workbook
            .edit_cell(SheetId(0), at, input, CalculationOptions::default())
            .unwrap();
        assert!(
            result
                .changed
                .iter()
                .any(|cell| cell.cell == CellRef::new(0, 0))
        );
        let after = warm_edit_reads(&workbook, at);
        assert_eq!(warm_edit_counters(), (0, 0, 0));
        assert_eq!(before, after);
        assert_eq!(
            workbook
                .sheet(SheetId(0))
                .unwrap()
                .cell(CellRef::new(0, 0))
                .unwrap()
                .value,
            CellValue::Number {
                value: input.parse().unwrap()
            }
        );
        assert_warm_edit_bounds(
            &workbook,
            Some(CellRange::new(CellRef::new(0, 0), CellRef::new(99_999, 10))),
        );
    }
}

#[test]
fn warm_edit_edge_shrinks_and_outside_extends() {
    let mut input = model(20, 11);
    let edge = CellRef::new(25, 15);
    input.sheets[0].set_cell(
        edge,
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::from_model(input).unwrap();
    let before = warm_edit_reads(&workbook, edge).0;
    reset_warm_edit_counters();
    workbook
        .edit_cell(SheetId(0), edge, "", CalculationOptions::default())
        .unwrap();
    let after = warm_edit_reads(&workbook, edge).0;
    assert_eq!(xlsx_model::workbook::used_range_scans(), 1);
    assert!(after.content_width < before.content_width);
    assert!(after.content_height < before.content_height);
    assert_warm_edit_bounds(
        &workbook,
        Some(CellRange::new(CellRef::new(0, 0), CellRef::new(19, 10))),
    );
    reset_warm_edit_counters();
    workbook
        .edit_cell(SheetId(0), edge, "2", CalculationOptions::default())
        .unwrap();
    assert_eq!(warm_edit_reads(&workbook, edge).0, before);
    assert_eq!(warm_edit_counters(), (0, 0, 0));
    assert_warm_edit_bounds(&workbook, Some(CellRange::new(CellRef::new(0, 0), edge)));
}

#[test]
fn warm_edit_style_and_structure_invalidate_bounds() {
    let mut input = model(20, 11);
    input.styles.fonts.push(Font {
        size_pt: Some(11.0),
        ..Font::default()
    });
    input.styles.cell_xfs.push(Xf {
        font: Some(0),
        ..Xf::default()
    });
    let mut workbook = Workbook::from_model(input).unwrap();
    let at = CellRef::new(2, 2);
    for (op, rebuilds) in [
        (
            Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::new(at, at),
                patch: StylePatch {
                    font_size: Some(32.0),
                    ..StylePatch::default()
                },
            },
            0,
        ),
        (
            Op::InsertRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            },
            1,
        ),
    ] {
        warm_edit_reads(&workbook, at);
        reset_warm_edit_counters();
        workbook
            .apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
        warm_edit_reads(&workbook, at);
        assert_eq!(xlsx_model::workbook::used_range_scans(), 1);
        assert_eq!(geometry_counters().0, rebuilds);
        let bounds = workbook.sheet(SheetId(0)).unwrap().used_range();
        assert_warm_edit_bounds(&workbook, bounds);
    }
}

#[test]
fn warm_edit_storage_counter_covers_build_clone_and_growth() {
    let input = model(20, 11);
    reset_warm_edit_counters();
    let mut geometry = GridGeometry::new(&input.sheets[0], &input.styles);
    assert_eq!(warm_edit_counters(), (1, 0, 1));
    let retained = geometry.clone();
    assert_eq!(warm_edit_counters(), (2, 0, 1));
    let mut sheet = input.sheets[0].clone();
    sheet.row_heights.insert(100_000, 20.0);
    geometry.update_rows(&sheet, &input.styles, [100_000]);
    assert_eq!(warm_edit_counters(), (3, 0, 1));
    assert_ne!(geometry.row_y(100_001), retained.row_y(100_001));
    sheet.used_range();
    assert_eq!(warm_edit_counters(), (3, 1, 1));
}

#[test]
fn warm_edit_spill_updates_bounds() {
    let mut input = model(1, 2);
    input.sheets[0].format.default_row_height_pt = Some(20.0);
    let anchor = CellRef::new(3, 0);
    input.sheets[0].set_cell(
        anchor,
        Cell {
            formula: Some("_xlfn.SEQUENCE(B1)".into()),
            ..Cell::default()
        },
    );
    input.sheets[0].set_array_formula(anchor, CellRange::new(anchor, anchor));
    let mut workbook = Workbook::from_model(input).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    warm_edit_reads(&workbook, anchor);
    for (value, end, scans) in [("4", 6, 1), ("2", 4, 1)] {
        reset_warm_edit_counters();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 1),
                value,
                CalculationOptions::default(),
            )
            .unwrap();
        warm_edit_reads(&workbook, anchor);
        assert_eq!(warm_edit_counters(), (0, scans, 0));
        assert_warm_edit_bounds(
            &workbook,
            Some(CellRange::new(CellRef::new(0, 0), CellRef::new(end, 1))),
        );
        assert_cold_geometry(&workbook, CellRef::new(7, 0));
    }
}

#[test]
fn warm_edit_collaborative_spill_updates_bounds() {
    let mut input = model(1, 2);
    input.sheets[0].format.default_row_height_pt = Some(20.0);
    let anchor = CellRef::new(3, 0);
    input.sheets[0].set_cell(
        anchor,
        Cell {
            formula: Some("_xlfn.SEQUENCE(B1)".into()),
            ..Cell::default()
        },
    );
    input.sheets[0].set_array_formula(anchor, CellRange::new(anchor, anchor));
    let mut workbook = Workbook::from_model_collaborative(input, 51).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    for (value, end) in [("4", 6), ("2", 4), ("4", 6)] {
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 1),
                value,
                CalculationOptions::default(),
            )
            .unwrap();
        let bounds = workbook.sheet(SheetId(0)).unwrap().used_range();
        assert_eq!(
            bounds,
            Some(CellRange::new(CellRef::new(0, 0), CellRef::new(end, 1)))
        );
        warm_edit_reads(&workbook, anchor);
        assert_warm_edit_bounds(&workbook, bounds);
        assert_cold_geometry(&workbook, CellRef::new(7, 0));
    }
}

#[test]
fn warm_edit_collaborative_cross_sheet_spill_updates_bounds() {
    let mut input = model(1, 2);
    input.sheets[0].name = "X".into();
    input.sheets[0].format.default_row_height_pt = Some(20.0);
    let anchor = CellRef::new(3, 0);
    input.sheets[0].set_cell(
        anchor,
        Cell {
            formula: Some("_xlfn.SEQUENCE(Y!B1)".into()),
            ..Cell::default()
        },
    );
    input.sheets[0].set_array_formula(anchor, CellRange::new(anchor, anchor));
    let mut other = Sheet::new("Y");
    other.set_cell(
        CellRef::new(0, 1),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    input.sheets.push(other);
    let mut workbook = Workbook::from_model_collaborative(input, 51).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(0, 1),
            "4",
            CalculationOptions::default(),
        )
        .unwrap();
    warm_edit_reads(&workbook, anchor);
    assert_cold_geometry(&workbook, CellRef::new(7, 0));
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(0, 1),
            "2",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_cold_geometry(&workbook, CellRef::new(7, 0));
}
