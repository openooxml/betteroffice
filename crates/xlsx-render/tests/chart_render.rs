use std::cell::Cell;

use ooxml_drawingml::chart::{ChartLegend, ChartPlotGroup, ChartSeries, ChartSpace};
use serde_json::{Value, json};
use xlsx_model::chart::{AnchorCell, AnchorExtent, AnchorPos, ChartAnchor, SheetChart};
use xlsx_model::workbook::{Cell as SheetCell, FreezePane, Sheet};
use xlsx_model::{CellRef, CellValue, SheetId, Workbook};
use xlsx_render::{
    DisplayList, DrawCmd, GridGeometry, RenderError, Viewport, build_display_list,
    build_display_list_with_charts,
};

fn chart_space() -> ChartSpace {
    ChartSpace {
        chart_type: "pie".into(),
        title: Some("Revenue".into()),
        legend: Some(ChartLegend {
            position: Some("right".into()),
            visible: false,
            ..Default::default()
        }),
        series: vec![ChartSeries {
            name: Some("Total".into()),
            categories: vec!["All".into()],
            values: vec![5.0],
            color: "#4472C4".into(),
            index: None,
            order: None,
            category_formula: None,
            value_formula: None,
            axis_ids: None,
            points: None,
            grouping: None,
            marker: None,
            smooth: None,
            ..Default::default()
        }],
        axes: None,
        plot_groups: Vec::new(),
        axis_list: None,
        ..Default::default()
    }
}

fn sheet_chart(anchor: ChartAnchor) -> SheetChart {
    SheetChart {
        part: "xl/charts/chart1.xml".into(),
        drawing: "xl/drawings/drawing1.xml".into(),
        anchor_index: 0,
        anchor,
        refs: Vec::new(),
    }
}

fn absolute_chart() -> SheetChart {
    sheet_chart(ChartAnchor::Absolute {
        pos: AnchorPos {
            x: 95_250,
            y: 190_500,
        },
        extent: AnchorExtent {
            cx: 1_905_000,
            cy: 1_143_000,
        },
    })
}

fn is_chart_command(command: &DrawCmd) -> bool {
    match command {
        DrawCmd::FillRect { clip, .. }
        | DrawCmd::Line { clip, .. }
        | DrawCmd::Path { clip, .. } => clip.is_some(),
        DrawCmd::Text { chart, .. } => *chart,
    }
}

fn snapshot_command(command: &DrawCmd) -> Value {
    match command {
        DrawCmd::Path {
            commands,
            fill,
            stroke,
            clip,
        } => json!({
            "op": "path",
            "commandCount": commands.len(),
            "first": commands.first(),
            "last": commands.last(),
            "fill": fill,
            "stroke": stroke,
            "clip": clip,
        }),
        command => serde_json::to_value(command).unwrap(),
    }
}

#[test]
fn chart_display_list_matches_snapshot() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.charts.push(absolute_chart());
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let display_list = build_display_list_with_charts(
        &workbook,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 300.0,
            height: 200.0,
        },
        |_| Ok(chart_space()),
    )
    .unwrap();
    let commands = display_list
        .commands
        .iter()
        .filter(|command| is_chart_command(command))
        .map(snapshot_command)
        .collect::<Vec<_>>();
    let actual = serde_json::to_string_pretty(&json!({
        "commands": commands,
        "charts": display_list.charts,
    }))
    .unwrap();
    assert_eq!(
        actual,
        include_str!("snapshots/chart_display_list.json").trim()
    );
}

#[test]
fn offscreen_charts_do_not_resolve_or_plot() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.charts.push(sheet_chart(ChartAnchor::Absolute {
        pos: AnchorPos {
            x: 9_525_000,
            y: 9_525_000,
        },
        extent: AnchorExtent {
            cx: 952_500,
            cy: 952_500,
        },
    }));
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let resolved = Cell::new(false);
    let display_list = build_display_list_with_charts(
        &workbook,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 100.0,
        },
        |_| {
            resolved.set(true);
            Ok(chart_space())
        },
    )
    .unwrap();

    assert!(!resolved.get());
    assert!(display_list.charts.is_empty());
    assert!(!display_list.commands.iter().any(is_chart_command));
}

#[test]
fn charts_clip_to_the_body_and_precede_pane_dividers() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(2, 2)));
    sheet.set_cell(
        CellRef::new(3, 3),
        SheetCell {
            value: CellValue::Text {
                value: "under chart".into(),
            },
            ..SheetCell::default()
        },
    );
    sheet.charts.push(sheet_chart(ChartAnchor::OneCell {
        from: AnchorCell {
            col: 3,
            col_off: 0,
            row: 3,
            row_off: 0,
        },
        extent: AnchorExtent {
            cx: 1_905_000,
            cy: 1_143_000,
        },
    }));
    let geometry = GridGeometry::new(&sheet);
    let frozen_x = geometry.col_x(1);
    let frozen_y = geometry.row_y(1);
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let display_list = build_display_list_with_charts(
        &workbook,
        SheetId(0),
        &Viewport {
            x: geometry.col_x(2),
            y: geometry.row_y(2),
            width: 400.0,
            height: 240.0,
        },
        |_| Ok(chart_space()),
    )
    .unwrap();

    let cell_text = display_list
        .commands
        .iter()
        .position(|command| {
            matches!(
                command,
                DrawCmd::Text {
                    text,
                    chart: false,
                    ..
                } if text == "under chart"
            )
        })
        .unwrap();
    let first_chart = display_list
        .commands
        .iter()
        .position(is_chart_command)
        .unwrap();
    let first_divider = display_list
        .commands
        .iter()
        .position(|command| {
            matches!(
                command,
                DrawCmd::Line {
                    width,
                    clip: None,
                    ..
                } if *width == 2.0
            )
        })
        .unwrap();

    assert!(cell_text < first_chart);
    assert!(first_chart < first_divider);
    for command in display_list
        .commands
        .iter()
        .filter(|command| is_chart_command(command))
    {
        let clip = match command {
            DrawCmd::FillRect {
                clip: Some(clip), ..
            }
            | DrawCmd::Line {
                clip: Some(clip), ..
            }
            | DrawCmd::Path {
                clip: Some(clip), ..
            }
            | DrawCmd::Text { clip, .. } => clip,
            _ => continue,
        };
        assert!(clip.x >= frozen_x);
        assert!(clip.y >= frozen_y);
    }
}

/// one neutral fill plus its four border lines.
const PLACEHOLDER_OPS: usize = 5;

fn placeholder_fills(display_list: &DisplayList) -> usize {
    display_list
        .commands
        .iter()
        .filter(|command| {
            matches!(
                command,
                DrawCmd::FillRect { color, clip: Some(_), .. } if color == "#f2f2f2"
            )
        })
        .count()
}

/// A worksheet carrying one chart plus a cell, so every assertion below is
/// about what survives when that chart cannot be drawn.
fn frame_with_one_chart<F>(resolver: F) -> DisplayList
where
    F: FnMut(&SheetChart) -> Result<ChartSpace, RenderError>,
{
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(0, 0),
        SheetCell {
            value: CellValue::Text {
                value: "beside the chart".into(),
            },
            ..SheetCell::default()
        },
    );
    sheet.charts.push(absolute_chart());
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    build_display_list_with_charts(
        &workbook,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 300.0,
            height: 200.0,
        },
        resolver,
    )
    .unwrap()
}

fn assert_degrades_locally(display_list: &DisplayList, label: &str) {
    assert!(display_list.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text { text, chart: false, .. } if text == "beside the chart"
    )));
    assert!(!display_list.grid.row_offsets.is_empty());
    assert_eq!(placeholder_fills(display_list), 1);
    // the placeholder is the chart layer's *entire* contribution: one fill plus
    // four border lines, and nothing the refused plot managed to emit.
    assert_eq!(
        display_list
            .commands
            .iter()
            .filter(|command| is_chart_command(command))
            .count(),
        PLACEHOLDER_OPS
    );
    assert_eq!(display_list.charts.len(), 1);
    assert!(display_list.charts[0].placeholder);
    assert_eq!(display_list.charts[0].label, label);
}

#[test]
fn an_unparseable_chart_degrades_to_a_placeholder() {
    let display_list = frame_with_one_chart(|chart| {
        Err(RenderError::ChartParseFailed {
            part: chart.part.clone(),
        })
    });
    assert_degrades_locally(&display_list, "Chart, not shown");
}

#[test]
fn a_missing_chart_part_degrades_to_a_placeholder() {
    let display_list = frame_with_one_chart(|chart| {
        Err(RenderError::ChartPartMissing {
            part: chart.part.clone(),
        })
    });
    assert_degrades_locally(&display_list, "Chart, not shown");
}

#[test]
fn an_unsupported_family_degrades_to_a_placeholder() {
    let display_list = frame_with_one_chart(|_| {
        Ok(ChartSpace {
            chart_type: "treemap".into(),
            title: Some("Spend".into()),
            ..chart_space()
        })
    });
    assert_degrades_locally(&display_list, "Spend, treemap chart, not shown");
}

#[test]
fn an_unsupported_feature_degrades_to_a_placeholder() {
    let display_list = frame_with_one_chart(|_| {
        Ok(ChartSpace {
            title: Some("Spend".into()),
            series: Vec::new(),
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some("column".into()),
                grouping: Some("percentStacked".into()),
                series: vec![ChartSeries {
                    name: Some("Total".into()),
                    categories: vec!["All".into()],
                    values: vec![5.0],
                    color: "#4472C4".into(),
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..chart_space()
        })
    });
    assert_degrades_locally(
        &display_list,
        "Spend, column chart, 1 series, 1 categories, not shown",
    );
}

#[test]
fn a_chart_anchored_off_the_grid_is_skipped_entirely() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(0, 0),
        SheetCell {
            value: CellValue::Text {
                value: "beside the chart".into(),
            },
            ..SheetCell::default()
        },
    );
    sheet.charts.push(sheet_chart(ChartAnchor::Absolute {
        pos: AnchorPos { x: 0, y: 0 },
        extent: AnchorExtent { cx: 0, cy: 0 },
    }));
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let display_list = build_display_list_with_charts(
        &workbook,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 300.0,
            height: 200.0,
        },
        |_| Ok(chart_space()),
    )
    .unwrap();

    assert!(display_list.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text { text, chart: false, .. } if text == "beside the chart"
    )));
    assert!(display_list.charts.is_empty());
    assert_eq!(placeholder_fills(&display_list), 0);
}

/// `build_display_list` supplies no chart source at all, so every visible chart
/// is a placeholder rather than a refused frame.
#[test]
fn the_resolverless_builder_places_holders_instead_of_failing() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(0, 0),
        SheetCell {
            value: CellValue::Text {
                value: "beside the chart".into(),
            },
            ..SheetCell::default()
        },
    );
    sheet.charts.push(absolute_chart());
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let display_list = build_display_list(
        &workbook,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 300.0,
            height: 200.0,
        },
    )
    .unwrap();
    assert_degrades_locally(&display_list, "Chart, not shown");
}

#[test]
fn a_drawable_chart_is_not_marked_as_a_placeholder() {
    let display_list = frame_with_one_chart(|_| Ok(chart_space()));
    assert_eq!(display_list.charts.len(), 1);
    assert!(!display_list.charts[0].placeholder);
    assert_eq!(placeholder_fills(&display_list), 0);
    assert!(
        display_list
            .commands
            .iter()
            .any(|command| matches!(command, DrawCmd::Path { .. }))
    );
}
