use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{Input, LayoutBlock};
use serde_json::{Value, json};

fn table(before: Option<u64>, after: Option<u64>, span: u64) -> Value {
    let border = json!({"style": "single", "width": 1, "color": "#000000"});
    let mut row = json!({
        "id": "row", "height": 24, "heightRule": "exact",
        "cells": [{
            "id": "cell", "colSpan": span, "blocks": [], "background": "#abcdef",
            "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
            "borders": {"top": border, "right": border, "bottom": border, "left": border}
        }]
    });
    if let Some(before) = before {
        row["gridBefore"] = json!(before);
    }
    if let Some(after) = after {
        row["gridAfter"] = json!(after);
    }
    json!({
        "kind": "table", "id": "table", "layoutMode": "fixed",
        "columnWidths": [20, 80, 120], "rows": [row]
    })
}

fn render(table: Value) -> (Value, Value, Value) {
    let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([table])).unwrap();
    let measures = measure_blocks(&mut blocks, 220.0, &MeasurementConfig::default()).unwrap();
    let measure = serde_json::to_value(&measures[0]).unwrap();
    let input = json!({
        "measured": [{"block": blocks[0], "measure": measure}],
        "options": {
            "pageSize": {"w": 320, "h": 300},
            "margins": {"top": 50, "right": 50, "bottom": 50, "left": 50}
        }
    });
    let mut pagination: Input = serde_json::from_value(input).unwrap();
    let layout = docx_layout::compute_layout_input(&mut pagination).unwrap();
    let resident = serde_json::to_value(
        docx_layout::build_display_list_value_from_resident(&pagination, &layout, "{}").unwrap(),
    )
    .unwrap();
    let mut input = serde_json::to_value(&pagination).unwrap();
    let layout = serde_json::to_value(layout).unwrap();
    input["layout"] = layout.clone();
    let display =
        serde_json::to_value(docx_layout::build_display_list_value(&input.to_string()).unwrap())
            .unwrap();
    assert_eq!(display, resident);
    (measure, layout, display)
}

fn assert_cell(display: &Value, column: u64, x: f64, width: f64) {
    let primitives = display["pages"][0]["primitives"].as_array().unwrap();
    let fills: Vec<_> = primitives
        .iter()
        .filter(|primitive| primitive["fill"] == "#abcdef")
        .collect();
    assert_eq!(fills.len(), 1);
    assert_eq!(fills[0]["x"], x);
    assert_eq!(fills[0]["w"], width);
    assert_eq!(fills[0]["cell"]["col"], column);
    let borders: Vec<_> = primitives
        .iter()
        .filter(|primitive| primitive["role"] == "table-border")
        .collect();
    assert_eq!(borders.len(), 4);
    for border in borders {
        for key in ["x1", "x2"] {
            let coordinate = border[key].as_f64().unwrap();
            assert!(coordinate >= x - 0.5 && coordinate <= x + width + 0.5);
        }
        assert_eq!(border["cell"]["col"], column);
    }
}

#[test]
fn grid_before_places_a_spanning_cell_in_columns_two_and_three() {
    let (measure, _, display) = render(table(Some(1), None, 2));
    assert_eq!(measure["columnWidths"], json!([20.0, 80.0, 120.0]));
    assert_eq!(measure["rows"][0]["cells"][0]["width"], 200.0);
    assert_cell(&display, 1, 70.0, 200.0);
}

#[test]
fn grid_after_leaves_the_last_column_unpainted() {
    let (measure, _, display) = render(table(None, Some(1), 2));
    assert_eq!(measure["columnWidths"], json!([20.0, 80.0, 120.0]));
    assert_eq!(measure["rows"][0]["cells"][0]["width"], 100.0);
    assert_cell(&display, 0, 50.0, 100.0);
}

#[test]
fn grid_before_and_after_paint_only_the_middle_column() {
    let (measure, _, display) = render(table(Some(1), Some(1), 1));
    assert_eq!(measure["rows"][0]["cells"].as_array().unwrap().len(), 1);
    assert_eq!(measure["rows"][0]["cells"][0]["width"], 80.0);
    assert_cell(&display, 1, 70.0, 80.0);
}

#[test]
fn omitted_grid_widths_use_dxa_and_table_relative_percentages() {
    let mut table = table(Some(1), Some(1), 1);
    table["rows"][0]["widthBefore"] = json!({"value": 600, "type": "dxa"});
    table["rows"][0]["widthAfter"] = json!({"value": 2500, "type": "pct"});
    let (measure, _, display) = render(table);
    assert_eq!(measure["columnWidths"], json!([40.0, 80.0, 110.0]));
    assert_cell(&display, 1, 90.0, 80.0);
}

#[test]
fn trailing_omitted_width_is_respected_without_a_leading_skip() {
    let mut table = table(None, Some(1), 2);
    table["rows"][0]["widthAfter"] = json!({"value": 900, "type": "dxa"});
    let (measure, _, display) = render(table);
    assert_eq!(measure["columnWidths"], json!([20.0, 80.0, 60.0]));
    assert_cell(&display, 0, 50.0, 100.0);
}

#[test]
fn zero_omitted_widths_keep_the_grid_positions_without_a_visible_gap() {
    let mut table = table(Some(1), Some(1), 1);
    table["rows"][0]["widthBefore"] = json!({"value": 0, "type": "dxa"});
    table["rows"][0]["widthAfter"] = json!({"value": 0, "type": "pct"});
    let (measure, _, display) = render(table);
    assert_eq!(measure["columnWidths"], json!([0.0, 80.0, 0.0]));
    assert_cell(&display, 1, 50.0, 80.0);
}

#[test]
fn bidi_rows_mirror_the_omitted_columns_with_their_cells() {
    let mut table = table(Some(1), Some(1), 1);
    table["bidi"] = json!(true);
    let (_, _, display) = render(table);
    assert_cell(&display, 1, 170.0, 80.0);
}

#[test]
fn rows_without_omitted_columns_keep_identical_layout_bytes() {
    let plain = render(table(None, None, 3));
    let zero = render(table(Some(0), Some(0), 3));
    assert_eq!(
        serde_json::to_string(&plain).unwrap(),
        serde_json::to_string(&zero).unwrap()
    );
    assert_eq!(plain.0["rows"][0]["cells"][0]["width"], 220.0);
    assert_cell(&plain.2, 0, 50.0, 220.0);
}
