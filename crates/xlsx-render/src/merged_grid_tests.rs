use super::*;
use xlsx_model::workbook::FreezePane;

fn workbook(merge: &str) -> Workbook {
    let mut sheet = Sheet::new("Merged");
    sheet.merges.push(CellRange::parse_a1(merge).unwrap());
    for index in 0..8 {
        sheet.col_widths.insert(index, 10.0);
        sheet.row_heights.insert(index, 15.0);
    }
    let mut wb = Workbook::default();
    wb.sheets.push(sheet);
    wb
}

fn viewport(geom: &GridGeometry) -> Viewport {
    Viewport {
        x: 0.0,
        y: 0.0,
        width: geom.col_x(7),
        height: geom.row_y(7),
    }
}

fn original_grid_commands(
    rows: &AxisLayout,
    cols: &AxisLayout,
    print: Option<(&PrintMetrics, bool)>,
) -> Vec<DrawCmd> {
    let gridline_color: Arc<str> = if print.is_some() {
        TEXT_COLOR
    } else {
        GRIDLINE_COLOR
    }
    .into();
    let mut grid_commands = Vec::new();
    let grid_offset = print.map_or(0.0, |(m, _)| 48.0 / m.dpi);
    let row_offsets = rows.offsets();
    let col_offsets = cols.offsets();
    let top = row_offsets.first().copied().unwrap_or(0.0);
    let bottom = row_offsets.last().copied().unwrap_or(0.0);
    let left = col_offsets.first().copied().unwrap_or(0.0);
    let right = col_offsets.last().copied().unwrap_or(0.0);
    for &x in &col_offsets {
        grid_commands.push(DrawCmd::Line {
            x1: x + grid_offset,
            y1: top + grid_offset,
            x2: x + grid_offset,
            y2: bottom + grid_offset,
            width: print.map_or(GRIDLINE_WIDTH, |(m, _)| 96.0 / m.dpi),
            color: gridline_color.clone(),
            style: None,
            clip: None,
        });
    }
    for &y in &row_offsets {
        grid_commands.push(DrawCmd::Line {
            x1: left + grid_offset,
            y1: y + grid_offset,
            x2: right + grid_offset,
            y2: y + grid_offset,
            width: print.map_or(GRIDLINE_WIDTH, |(m, _)| 96.0 / m.dpi),
            color: gridline_color.clone(),
            style: None,
            clip: None,
        });
    }
    grid_commands
}

fn frame_lines(wb: &Workbook, vp: &Viewport, print: Option<(&PrintMetrics, bool)>) -> Vec<DrawCmd> {
    build_frame(
        wb,
        SheetId(0),
        vp,
        &[],
        |chart| {
            Err::<ChartSpace, _>(RenderError::ChartSourceUnavailable {
                part: chart.part.clone(),
            })
        },
        print,
    )
    .unwrap()
    .commands
    .into_iter()
    .filter(|command| matches!(command, DrawCmd::Line { .. }))
    .collect()
}

fn append_frame_edges(
    commands: &mut Vec<DrawCmd>,
    rows: &AxisLayout,
    cols: &AxisLayout,
    vp: &Viewport,
    print: Option<(&PrintMetrics, bool)>,
) {
    if let Some((metrics, true)) = print {
        let width = 96.0 / metrics.dpi;
        let offset = width / 2.0;
        for (x1, y1, x2, y2) in [
            (offset, offset, vp.width + offset, offset),
            (offset, offset, offset, vp.height + offset),
            (
                offset,
                vp.height + offset,
                vp.width + offset,
                vp.height + offset,
            ),
            (
                vp.width + offset,
                offset,
                vp.width + offset,
                vp.height + offset,
            ),
        ] {
            commands.push(DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color: TEXT_COLOR.into(),
                style: None,
                clip: None,
            });
        }
    }
    if let Some(x) = cols.divider {
        commands.push(DrawCmd::Line {
            x1: x,
            y1: 0.0,
            x2: x,
            y2: vp.height,
            width: PANE_DIVIDER_WIDTH,
            color: PANE_DIVIDER_COLOR.into(),
            style: None,
            clip: None,
        });
    }
    if let Some(y) = rows.divider {
        commands.push(DrawCmd::Line {
            x1: 0.0,
            y1: y,
            x2: vp.width,
            y2: y,
            width: PANE_DIVIDER_WIDTH,
            color: PANE_DIVIDER_COLOR.into(),
            style: None,
            clip: None,
        });
    }
}

fn print_metrics(dpi: f32) -> PrintMetrics {
    PrintMetrics {
        dpi,
        max_digit_width: 6.0,
        default_row_height_pt: 15.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    }
}

fn assert_original_grid(wb: &Workbook, vp: &Viewport, print: Option<(&PrintMetrics, bool)>) {
    let geom = print.map_or_else(
        || GridGeometry::new(&wb.sheets[0], &wb.styles),
        |(metrics, _)| GridGeometry::for_print(&wb.sheets[0], &wb.styles, metrics),
    );
    let (rows, cols, _, _) = viewport_axes(&wb.sheets[0], vp, &geom, print.is_some());
    let mut expected = if print.is_some_and(|(_, gridlines)| !gridlines) {
        Vec::new()
    } else {
        original_grid_commands(&rows, &cols, print)
    };
    append_frame_edges(&mut expected, &rows, &cols, vp, print);
    assert_eq!(frame_lines(wb, vp, print), expected);
}

#[test]
fn no_merges_match_original_grid_order_extents_and_multiplicity() {
    for (frozen_rows, frozen_cols) in [(0, 0), (2, 0), (0, 2), (2, 2)] {
        for hidden in [false, true] {
            let mut wb = workbook("B2:D4");
            wb.sheets[0].merges.clear();
            if frozen_rows > 0 || frozen_cols > 0 {
                wb.sheets[0].freeze_pane = Some(FreezePane::new(
                    frozen_rows,
                    frozen_cols,
                    CellRef::new(4, 4),
                ));
            }
            if hidden {
                wb.sheets[0].col_widths.insert(5, 0.0);
                wb.sheets[0].row_heights.insert(5, 0.0);
            }
            let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
            for (x, y) in [(0.25, 0.25), (geom.col_x(3) + 0.25, geom.row_y(3) + 0.25)] {
                let vp = Viewport {
                    x,
                    y,
                    width: geom.col_x(7) - 0.5,
                    height: geom.row_y(7) - 0.5,
                };
                assert_original_grid(&wb, &vp, None);
                for (width, height) in [(0.0, vp.height), (vp.width, 0.0), (0.0, 0.0)] {
                    assert_original_grid(
                        &wb,
                        &Viewport {
                            width,
                            height,
                            ..vp
                        },
                        None,
                    );
                }
            }
        }
    }
}

#[test]
fn no_merges_match_original_print_grid_at_two_dpis() {
    let mut wb = workbook("B2:D4");
    wb.sheets[0].merges.clear();
    wb.sheets[0].freeze_pane = Some(FreezePane::new(2, 2, CellRef::new(4, 4)));
    wb.sheets[0].col_widths.insert(3, 0.0);
    wb.sheets[0].row_heights.insert(1, 0.0);
    for dpi in [72.0, 144.0] {
        let metrics = print_metrics(dpi);
        for gridlines in [false, true] {
            let vp = Viewport {
                x: 0.25,
                y: 0.25,
                width: 400.0,
                height: 40.0,
            };
            assert_original_grid(&wb, &vp, Some((&metrics, gridlines)));
            for (width, height) in [(0.0, vp.height), (vp.width, 0.0), (0.0, 0.0)] {
                assert_original_grid(
                    &wb,
                    &Viewport {
                        width,
                        height,
                        ..vp
                    },
                    Some((&metrics, gridlines)),
                );
            }
        }
    }
}

fn assert_merge_subtraction(
    wb: &Workbook,
    vp: &Viewport,
    geom: &GridGeometry,
    print: Option<(&PrintMetrics, bool)>,
    interior: [(f32, f32); 2],
) {
    let (rows, cols, _, _) = viewport_axes(&wb.sheets[0], vp, geom, print.is_some());
    let reference = original_grid_commands(&rows, &cols, print);
    let col_boundaries: Vec<_> = std::iter::once(cols.tracks[0].index)
        .chain(cols.tracks.iter().map(|track| track.index + 1))
        .collect();
    let row_boundaries = std::iter::once(rows.tracks[0].index)
        .chain(rows.tracks.iter().map(|track| track.index + 1));
    let boundaries = col_boundaries
        .iter()
        .map(|&index| (true, index))
        .chain(row_boundaries.map(|index| (false, index)));
    let offset = print.map_or(0.0, |(metrics, _)| 48.0 / metrics.dpi);
    let mut expected = Vec::new();
    for (command, (vertical, index)) in reference.into_iter().zip(boundaries) {
        let position = match &command {
            DrawCmd::Line { x1, y1, .. } => {
                if vertical {
                    *x1
                } else {
                    *y1
                }
            }
            _ => unreachable!(),
        };
        let (axis_start, axis_end) = interior[usize::from(!vertical)];
        if !(2..=3).contains(&index)
            || position <= axis_start + offset
            || position >= axis_end + offset
        {
            expected.push(command);
            continue;
        }
        let (lower, upper) = interior[usize::from(vertical)];
        let (lower, upper) = (lower + offset, upper + offset);
        let mut before = command.clone();
        let mut after = command;
        if let DrawCmd::Line { x1, y1, x2, y2, .. } = &mut before {
            let start = if vertical { *y1 } else { *x1 };
            if lower > start {
                if vertical {
                    *y2 = lower;
                } else {
                    *x2 = lower;
                }
                expected.push(before);
            }
        }
        if let DrawCmd::Line { x1, y1, x2, y2, .. } = &mut after {
            let end = if vertical { *y2 } else { *x2 };
            if upper < end {
                if vertical {
                    *y1 = upper;
                } else {
                    *x1 = upper;
                }
                expected.push(after);
            }
        }
    }
    append_frame_edges(&mut expected, &rows, &cols, vp, print);
    assert_eq!(frame_lines(wb, vp, print), expected);
}

#[test]
fn merge_changes_only_crossing_strokes_by_exact_interval_subtraction() {
    for frozen in [false, true] {
        let mut wb = workbook("B2:D4");
        wb.sheets[0].col_widths.insert(2, 0.0);
        wb.sheets[0].row_heights.insert(2, 0.0);
        if frozen {
            wb.sheets[0].freeze_pane = Some(FreezePane::new(2, 2, CellRef::new(2, 2)));
        }
        let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
        for (x, y) in [(0.25, 0.25), (geom.col_x(2) + 0.25, geom.row_y(2) + 0.25)]
            .into_iter()
            .take(if frozen { 1 } else { 2 })
        {
            let vp = Viewport {
                x,
                y,
                ..viewport(&geom)
            };
            let left = geom.col_x(1) - if frozen { 0.0 } else { vp.x };
            let top = geom.row_y(1) - if frozen { 0.0 } else { vp.y };
            assert_merge_subtraction(
                &wb,
                &vp,
                &geom,
                None,
                [(left, geom.col_x(4) - vp.x), (top, geom.row_y(4) - vp.y)],
            );
        }
    }
    let wb = workbook("B2:D4");
    for dpi in [72.0, 144.0] {
        let metrics = print_metrics(dpi);
        let geom = GridGeometry::for_print(&wb.sheets[0], &wb.styles, &metrics);
        let vp = Viewport {
            x: 0.25,
            y: 0.25,
            width: geom.col_x(5) - 0.5,
            height: geom.row_y(5) - 0.5,
        };
        assert_merge_subtraction(
            &wb,
            &vp,
            &geom,
            Some((&metrics, true)),
            [
                (geom.col_x(1) - vp.x, geom.col_x(4) - vp.x),
                (geom.row_y(1) - vp.y, geom.row_y(4) - vp.y),
            ],
        );
        assert!(frame_lines(&wb, &vp, Some((&metrics, false))).is_empty());
    }
}

fn line_covers(dl: &DisplayList, x: f32, y: f32) -> bool {
    dl.commands.iter().any(|cmd| match cmd {
        DrawCmd::Line { x1, y1, x2, y2, .. } => {
            let vertical = (x1 - x2).abs() < 0.01 && (x - x1).abs() < 0.01;
            let horizontal = (y1 - y2).abs() < 0.01 && (y - y1).abs() < 0.01;
            (vertical || horizontal)
                && x >= x1.min(*x2) - 0.01
                && x <= x1.max(*x2) + 0.01
                && y >= y1.min(*y2) - 0.01
                && y <= y1.max(*y2) + 0.01
        }
        _ => false,
    })
}

fn assert_merge_edges(dl: &DisplayList, left: f32, top: f32, right: f32, bottom: f32) {
    let x = (left + right) / 2.0;
    let y = (top + bottom) / 2.0;
    for (x, y) in [(left, y), (right, y), (x, top), (x, bottom)] {
        assert!(line_covers(dl, x, y), "missing merge edge at ({x}, {y})");
    }
}

#[test]
fn empty_merge_preserves_outer_edges_and_grid_continuations() {
    let wb = workbook("B2:C3");
    let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
    let dl = build_display_list(&wb, SheetId(0), &viewport(&geom)).unwrap();
    for row in [1.25, 1.75, 2.25, 2.75] {
        assert!(!line_covers(&dl, geom.col_x(2), geom.row_y(1) * row));
    }
    for col in [1.25, 1.75, 2.25, 2.75] {
        assert!(!line_covers(&dl, geom.col_x(1) * col, geom.row_y(2)));
    }
    assert_merge_edges(
        &dl,
        geom.col_x(1),
        geom.row_y(1),
        geom.col_x(3),
        geom.row_y(3),
    );
    for (x, y) in [
        (geom.col_x(2), geom.row_y(1) / 2.0),
        (geom.col_x(2), (geom.row_y(3) + geom.row_y(4)) / 2.0),
        (geom.col_x(1) / 2.0, geom.row_y(2)),
        ((geom.col_x(3) + geom.col_x(4)) / 2.0, geom.row_y(2)),
    ] {
        assert!(line_covers(&dl, x, y));
    }
}

#[test]
fn adjacent_merges_union_their_blocked_grid_intervals() {
    let mut wb = workbook("B2:C3");
    wb.sheets[0]
        .merges
        .push(CellRange::parse_a1("B4:C5").unwrap());
    let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
    let dl = build_display_list(&wb, SheetId(0), &viewport(&geom)).unwrap();
    for row in 1..5 {
        let y = (geom.row_y(row) + geom.row_y(row + 1)) / 2.0;
        assert!(!line_covers(&dl, geom.col_x(2), y));
    }
    assert!(!line_covers(&dl, geom.col_x(2), geom.row_y(3) - 0.1));
    assert!(!line_covers(&dl, geom.col_x(2), geom.row_y(3) + 0.1));
    assert!(line_covers(&dl, geom.col_x(2), geom.row_y(1) / 2.0));
    assert!(line_covers(
        &dl,
        geom.col_x(2),
        (geom.row_y(5) + geom.row_y(6)) / 2.0,
    ));
    assert!(line_covers(
        &dl,
        (geom.col_x(1) + geom.col_x(2)) / 2.0,
        geom.row_y(3),
    ));
}

#[test]
fn scrolled_out_merge_anchor_suppresses_grid_with_fractional_scroll() {
    let wb = workbook("B2:D4");
    let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
    let mut vp = viewport(&geom);
    vp.x = geom.col_x(2) + 0.25;
    vp.y = geom.row_y(2) + 0.75;
    let dl = build_display_list(&wb, SheetId(0), &vp).unwrap();
    let x = geom.col_x(3) - vp.x;
    let y = geom.row_y(3) - vp.y;
    assert!(!line_covers(&dl, x, y / 2.0));
    assert!(!line_covers(&dl, x / 2.0, y));
    assert!(line_covers(
        &dl,
        x,
        (geom.row_y(4) + geom.row_y(5)) / 2.0 - vp.y,
    ));
    assert!(line_covers(
        &dl,
        (geom.col_x(4) + geom.col_x(5)) / 2.0 - vp.x,
        y,
    ));
    assert!(line_covers(&dl, geom.col_x(4) - vp.x, y / 2.0));
    assert!(line_covers(&dl, x / 2.0, geom.row_y(4) - vp.y));
}

#[test]
fn merge_crossing_frozen_splits_suppresses_each_pane_band() {
    let mut wb = workbook("A1:F6");
    wb.sheets[0].freeze_pane = Some(FreezePane::new(2, 2, CellRef::new(4, 4)));
    let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
    let mut vp = viewport(&geom);
    vp.x = geom.col_x(1) * 2.25;
    vp.y = geom.row_y(1) * 2.25;
    let dl = build_display_list(&wb, SheetId(0), &vp).unwrap();
    let body_x = geom.col_x(5) - vp.x;
    let body_y = geom.row_y(5) - vp.y;
    for x in [geom.col_x(1), body_x] {
        for y in [
            geom.row_y(1) / 2.0,
            (geom.row_y(5) + geom.row_y(6)) / 2.0 - vp.y,
        ] {
            assert!(!line_covers(&dl, x, y));
        }
    }
    for y in [geom.row_y(1), body_y] {
        for x in [
            geom.col_x(1) / 2.0,
            (geom.col_x(5) + geom.col_x(6)) / 2.0 - vp.x,
        ] {
            assert!(!line_covers(&dl, x, y));
        }
    }
    assert!(line_covers(&dl, geom.col_x(2), geom.row_y(1) / 2.0));
    assert!(line_covers(&dl, geom.col_x(1) / 2.0, geom.row_y(2)));
    assert!(line_covers(
        &dl,
        body_x,
        (geom.row_y(6) + geom.row_y(7)) / 2.0 - vp.y,
    ));
    assert!(line_covers(
        &dl,
        (geom.col_x(6) + geom.col_x(7)) / 2.0 - vp.x,
        body_y,
    ));
    for (x, y) in [
        (0.0, (geom.row_y(5) + geom.row_y(6)) / 2.0 - vp.y),
        (
            geom.col_x(6) - vp.x,
            (geom.row_y(5) + geom.row_y(6)) / 2.0 - vp.y,
        ),
        ((geom.col_x(5) + geom.col_x(6)) / 2.0 - vp.x, 0.0),
        (
            (geom.col_x(5) + geom.col_x(6)) / 2.0 - vp.x,
            geom.row_y(6) - vp.y,
        ),
    ] {
        assert!(line_covers(&dl, x, y));
    }
}

#[test]
fn merge_spanning_hidden_tracks_preserves_collapsed_outer_edges() {
    for hidden in 1..=3 {
        let mut wb = workbook("B2:D4");
        wb.sheets[0].col_widths.insert(hidden, 0.0);
        wb.sheets[0].row_heights.insert(hidden, 0.0);
        let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
        let dl = build_display_list(&wb, SheetId(0), &viewport(&geom)).unwrap();
        let interior = if hidden == 1 { 3 } else { 2 };
        let x = (geom.col_x(1) + geom.col_x(4)) / 2.0;
        let y = (geom.row_y(1) + geom.row_y(4)) / 2.0;
        assert!(!line_covers(&dl, geom.col_x(interior), y));
        assert!(!line_covers(&dl, x, geom.row_y(interior)));
        assert_merge_edges(
            &dl,
            geom.col_x(1),
            geom.row_y(1),
            geom.col_x(4),
            geom.row_y(4),
        );
        assert!(dl.commands.iter().all(|cmd| match cmd {
            DrawCmd::Line { x1, y1, x2, y2, .. } => x1 != x2 || y1 != y2,
            _ => true,
        }));
    }
}

#[test]
fn hidden_leading_and_trailing_merge_tracks_preserve_exact_grid_commands() {
    for (hidden_col, hidden_row) in [
        (Some(1), None),
        (Some(3), None),
        (None, Some(1)),
        (None, Some(3)),
    ] {
        for frozen in [0, 2, 4] {
            let mut wb = workbook("B2:D4");
            if let Some(col) = hidden_col {
                wb.sheets[0].col_widths.insert(col, 0.0);
            }
            if let Some(row) = hidden_row {
                wb.sheets[0].row_heights.insert(row, 0.0);
            }
            if frozen > 0 {
                wb.sheets[0].freeze_pane =
                    Some(FreezePane::new(frozen, frozen, CellRef::new(4, 4)));
            }
            let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
            for scroll in [0.0, 0.25] {
                let vp = Viewport {
                    x: scroll,
                    y: scroll,
                    ..viewport(&geom)
                };
                let leading_scroll = if frozen > 1 { 0.0 } else { scroll };
                let trailing_scroll = if frozen >= 4 { 0.0 } else { scroll };
                assert_merge_subtraction(
                    &wb,
                    &vp,
                    &geom,
                    None,
                    [
                        (
                            geom.col_x(1) - leading_scroll,
                            geom.col_x(4) - trailing_scroll,
                        ),
                        (
                            geom.row_y(1) - leading_scroll,
                            geom.row_y(4) - trailing_scroll,
                        ),
                    ],
                );
            }
        }
    }
}

#[test]
fn print_grid_suppresses_merge_interiors_and_keeps_region_perimeter() {
    let wb = workbook("B2:D4");
    let metrics = PrintMetrics {
        dpi: 72.0,
        max_digit_width: 6.0,
        default_row_height_pt: 15.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    };
    let geom = GridGeometry::for_print(&wb.sheets[0], &wb.styles, &metrics);
    for partial in [false, true] {
        let mut vp = viewport(&geom);
        if partial {
            vp.x = geom.col_x(2) + 0.25;
            vp.y = geom.row_y(2) + 0.75;
        }
        let dl =
            build_print_display_list_with_charts(&wb, SheetId(0), &vp, &metrics, true, |chart| {
                Err::<ChartSpace, _>(RenderError::ChartSourceUnavailable {
                    part: chart.part.clone(),
                })
            })
            .unwrap();
        let offset = 48.0 / metrics.dpi;
        let x = geom.col_x(3) - vp.x + offset;
        let y = geom.row_y(3) - vp.y + offset;
        assert!(!line_covers(
            &dl,
            x,
            (geom.row_y(2) + geom.row_y(3)) / 2.0 - vp.y + offset,
        ));
        assert!(!line_covers(
            &dl,
            (geom.col_x(2) + geom.col_x(3)) / 2.0 - vp.x + offset,
            y,
        ));
        if !partial {
            assert_merge_edges(
                &dl,
                geom.col_x(1) + offset,
                geom.row_y(1) + offset,
                geom.col_x(4) + offset,
                geom.row_y(4) + offset,
            );
        }
        assert_merge_edges(&dl, offset, offset, vp.width + offset, vp.height + offset);
        assert!(
            dl.commands
                .iter()
                .filter_map(|cmd| match cmd {
                    DrawCmd::Line { width, .. } => Some(*width),
                    _ => None,
                })
                .all(|width| (width - 96.0 / metrics.dpi).abs() < 0.01)
        );
    }
}
