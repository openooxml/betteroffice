use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{BlockExtent, LayoutBlock};
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

/// Lines the label takes in an autofit table whose stated widths overflow
/// the text width, with `slack` px of room beside the label.
fn label_lines(config: &MeasurementConfig, label_width: f64, slack: f64) -> usize {
    let twips = ((label_width + slack + 2.0 * PADDING) * 15.0).round();
    let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([{
        "kind": "table", "id": "table", "tableLayout": "autofit",
        "width": 0, "widthType": "auto",
        "gridWidths": [twips / 15.0, 150.0],
        "rows": [{"id": "row", "cells": [cell("label", twips, LABEL), cell("value", 2250.0, "x")]}]
    }]))
    .unwrap();
    let measures = measure_blocks(&mut blocks, 200.0, config).unwrap();
    let BlockExtent::Table(table) = &measures[0] else {
        panic!("table extent")
    };
    let BlockExtent::Paragraph(label) = &table.rows[0].cells[0].blocks[0] else {
        panic!("label extent")
    };
    label.lines.len()
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
    assert_eq!(label_lines(&config, label_width, 0.5), 1);
    assert_eq!(label_lines(&config, label_width, -0.5), 2);
}
