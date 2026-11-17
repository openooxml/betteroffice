use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::table_grid::resolve_table_column_widths;
use docx_layout::types::{BlockExtent, LayoutBlock, TableBlock, TableExtent};
use serde_json::{Value, json};

const LABEL: &str = "Requester/Trainer:";
const PADDING: f64 = 7.2;

fn config() -> MeasurementConfig {
    let font = docx_layout::register_measure_font(include_bytes!(
        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
    ))
    .unwrap();
    MeasurementConfig {
        font_chains: std::collections::BTreeMap::from([(
            "liberation sans|0|0".to_owned(),
            vec![font],
        )]),
        defaults: json!({"fontFamily": "Liberation Sans", "fontSize": 10}),
        ..MeasurementConfig::default()
    }
}

fn paragraph(text: &str) -> Value {
    json!({"kind": "paragraph", "id": text, "attrs": {"alignment": "right"},
        "runs": [{"kind": "text", "text": text, "fontFamily": "Liberation Sans", "fontSize": 10}]})
}

fn cell(id: &str, twips: f64, text: &str) -> Value {
    json!({"id": id, "preferredWidth": {"value": twips, "type": "dxa"},
        "padding": {"top": 0, "bottom": 0, "left": PADDING, "right": PADDING},
        "blocks": [paragraph(text)]})
}

fn measure_table(
    config: &MeasurementConfig,
    grid: [f64; 2],
    cells: [Value; 2],
    content_width: f64,
) -> TableExtent {
    let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
        "kind": "table", "id": "table", "tableLayout": "autofit",
        "width": 0, "widthType": "auto",
        "gridWidths": grid,
        "rows": [{"id": "row", "cells": cells}]
    }]))
    .unwrap();
    let measures = measure_blocks(&mut blocks, content_width, config).unwrap();
    let BlockExtent::Table(table) = &measures[0] else {
        panic!("table extent")
    };
    table.clone()
}

/// An autofit table whose stated widths overflow the text width, with `slack`
/// px of room beside the label, and the label column's stated width.
fn label_table(config: &MeasurementConfig, label_width: f64, slack: f64) -> (f64, TableExtent) {
    let twips = ((label_width + slack + 2.0 * PADDING) * 15.0).round();
    let table = measure_table(
        config,
        [twips / 15.0, 150.0],
        [cell("label", twips, LABEL), cell("value", 2250.0, "x")],
        200.0,
    );
    (twips / 15.0, table)
}

fn label_lines(table: &TableExtent) -> usize {
    let BlockExtent::Paragraph(label) = &table.rows[0].cells[0].blocks[0] else {
        panic!("label extent")
    };
    label.lines.len()
}

fn assert_widths(actual: &[f64], expected: &[f64]) {
    assert_eq!(actual.len(), expected.len(), "{actual:?} vs {expected:?}");
    for (actual_width, expected_width) in actual.iter().zip(expected) {
        assert!(
            (actual_width - expected_width).abs() < 1e-6,
            "{actual:?} vs {expected:?}"
        );
    }
}

#[test]
fn a_label_that_fits_its_stated_cell_width_stays_on_one_line() {
    let config = config();
    let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([paragraph(LABEL)])).unwrap();
    let BlockExtent::Paragraph(natural) = &measure_blocks(&mut blocks, 1000.0, &config).unwrap()[0]
    else {
        panic!("paragraph extent")
    };
    let label_width = natural.lines[0].width;
    let (stated, fitting) = label_table(&config, label_width, 0.5);
    assert_eq!(label_lines(&fitting), 1);
    assert_widths(&fitting.column_widths, &[stated, 150.0]);
    assert!(fitting.column_widths.iter().sum::<f64>() > 200.0);
    let (_, short) = label_table(&config, label_width, -0.5);
    assert_eq!(label_lines(&short), 2);
}

#[test]
fn content_minimums_past_the_budget_do_not_widen_stated_cells_that_fit() {
    let config = config();
    let wide = |id: &str| {
        let mut cell = cell(id, 1440.0, "x");
        cell["minContentWidth"] = json!(288.0);
        cell["maxContentWidth"] = json!(288.0);
        cell
    };
    let table = measure_table(&config, [96.0, 96.0], [wide("a"), wide("b")], 192.0);
    assert_widths(&table.column_widths, &[96.0, 96.0]);
}

#[test]
fn content_minimums_past_the_budget_do_not_widen_stated_cells_inside_a_wider_grid() {
    let config = config();
    let wide = |id: &str| {
        let mut cell = cell(id, 1440.0, "x");
        cell["minContentWidth"] = json!(288.0);
        cell["maxContentWidth"] = json!(288.0);
        cell
    };
    let table = measure_table(&config, [300.0, 300.0], [wide("a"), wide("b")], 192.0);
    assert_widths(&table.column_widths, &[300.0, 300.0]);
}

#[test]
fn content_minimum_past_the_limit_keeps_the_earlier_widths() {
    let wide = |id: u32| {
        json!({"id": id, "blocks": [],
            "preferredWidth": {"value": 3750.0, "type": "dxa"},
            "minContentWidth": 200.0, "maxContentWidth": 200.0})
    };
    let table: TableBlock = serde_json::from_value(json!({
        "id": 0, "layoutMode": "autofit",
        "gridWidths": [150.0, 150.0], "columnWidths": [100.0, 100.0],
        "rows": [{"id": 0, "cells": [wide(0), wide(1)]}]
    }))
    .unwrap();
    assert_widths(&resolve_table_column_widths(&table, 200.0), &[100.0, 100.0]);
}

#[test]
fn stated_widths_past_the_limit_with_unshrinkable_content_keep_the_earlier_widths() {
    let wide = |id: u32| {
        json!({"id": id, "blocks": [], "widthValue": 3000.0, "widthType": "dxa",
            "minContentWidth": 150.0, "maxContentWidth": 150.0})
    };
    let table: TableBlock = serde_json::from_value(json!({
        "id": 0, "layoutMode": "autofit",
        "gridWidths": [240.0],
        "rows": [{"id": 0, "cells": [wide(0), wide(1)]}]
    }))
    .unwrap();
    assert_widths(&resolve_table_column_widths(&table, 200.0), &[240.0, 240.0]);
}

#[test]
fn stated_widths_past_a_grid_that_fills_the_text_width_stay_inside_it() {
    let config = config();
    let table = measure_table(
        &config,
        [100.0, 100.0],
        [cell("a", 1800.0, "x"), cell("b", 1800.0, "x")],
        200.0,
    );
    assert_widths(&table.column_widths, &[100.0, 100.0]);
}
