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

fn line_covers(dl: &DisplayList, x: f32, y: f32) -> bool {
    dl.commands.iter().any(|cmd| match cmd {
        DrawCmd::Line {
            x1, y1, x2, y2, ..
        } => {
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
        (geom.col_x(6) - vp.x, (geom.row_y(5) + geom.row_y(6)) / 2.0 - vp.y),
        ((geom.col_x(5) + geom.col_x(6)) / 2.0 - vp.x, 0.0),
        ((geom.col_x(5) + geom.col_x(6)) / 2.0 - vp.x, geom.row_y(6) - vp.y),
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
            DrawCmd::Line {
                x1, y1, x2, y2, ..
            } => x1 != x2 || y1 != y2,
            _ => true,
        }));
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
        let dl = build_print_display_list_with_charts(
            &wb,
            SheetId(0),
            &vp,
            &metrics,
            true,
            |chart| Err::<ChartSpace, _>(RenderError::ChartSourceUnavailable {
                part: chart.part.clone(),
            }),
        )
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
