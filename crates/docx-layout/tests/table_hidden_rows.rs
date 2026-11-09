use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::table_grid::drop_hidden_rows;
use docx_layout::types::{Input, LayoutBlock, TableBlock};
use serde_json::{Value, json};

fn cell(id: &str, fill: &str) -> Value {
    let border = json!({"style": "single", "width": 2, "color": "#000000"});
    json!({
        "id": id, "blocks": [], "background": fill,
        "padding": {"top": 0, "bottom": 0, "left": 0, "right": 0},
        "borders": {"top": border, "right": border, "bottom": border, "left": border}
    })
}

fn row(id: &str, fill: &str, height: f64) -> Value {
    json!({"id": id, "height": height, "heightRule": "atLeast", "cells": [cell(&format!("{id}c"), fill)]})
}

fn table(rows: Vec<Value>) -> TableBlock {
    serde_json::from_value(json!({
        "kind": "table", "id": "table", "layoutMode": "fixed",
        "columnWidths": [120], "rows": rows
    }))
    .unwrap()
}

fn three_rows(hidden: Option<bool>) -> TableBlock {
    let mut middle = row("r1", "#ff0000", 4.0);
    if let Some(hidden) = hidden {
        middle["hidden"] = json!(hidden);
    }
    table(vec![
        row("r0", "#00aa00", 24.0),
        middle,
        row("r2", "#0000aa", 30.0),
    ])
}

fn render(table: TableBlock) -> (Value, Value) {
    let mut blocks = vec![LayoutBlock::Table(table)];
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
    let display = serde_json::to_value(
        docx_layout::build_display_list_value_from_resident(&pagination, &layout, "{}").unwrap(),
    )
    .unwrap();
    (measure, display)
}

fn primitives(display: &Value) -> Vec<Value> {
    display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().cloned().unwrap_or_default())
        .collect()
}

#[test]
fn a_hidden_middle_row_lays_out_as_if_absent() {
    let mut hidden = three_rows(Some(true));
    drop_hidden_rows(&mut hidden);
    let without = table(vec![row("r0", "#00aa00", 24.0), row("r2", "#0000aa", 30.0)]);
    let sources: Vec<_> = hidden.rows.iter().map(|row| row.source_row).collect();
    assert_eq!(sources, vec![Some(0), Some(2)]);
    let mut renumbered = hidden.clone();
    for row in &mut renumbered.rows {
        row.source_row = None;
    }
    assert_eq!(renumbered, without);
    let (_, renumbered_display) = render(renumbered);
    let (expected_measure, expected_display) = render(without);
    assert_eq!(renumbered_display, expected_display);

    let (measure, display) = render(hidden);
    assert_eq!(measure, expected_measure);
    let mut cell_rows: Vec<u64> = primitives(&display)
        .iter()
        .filter_map(|p| p["cell"]["row"].as_u64())
        .collect();
    cell_rows.sort_unstable();
    cell_rows.dedup();
    assert_eq!(cell_rows, vec![0, 2]);
    let row_counts: Vec<u64> = primitives(&display)
        .iter()
        .filter_map(|p| p["table"]["rowCount"].as_u64())
        .collect();
    assert!(!row_counts.is_empty() && row_counts.iter().all(|count| *count == 3));
    assert_eq!(measure["rows"].as_array().unwrap().len(), 2);
    let (painted_measure, _) = render(three_rows(None));
    assert!(
        measure["totalHeight"].as_f64().unwrap() < painted_measure["totalHeight"].as_f64().unwrap()
    );
    let primitives = primitives(&display);
    assert!(primitives.iter().all(|p| p["fill"] != "#ff0000"));
    assert!(primitives.iter().any(|p| p["role"] == "table-border"));
}

#[test]
fn layout_alone_still_paints_a_row_marked_hidden() {
    let (marked_measure, marked_display) = render(three_rows(Some(true)));
    let (plain_measure, plain_display) = render(three_rows(None));
    assert_eq!(marked_measure, plain_measure);
    assert_eq!(
        primitives(&marked_display).len(),
        primitives(&plain_display).len()
    );
    assert!(
        primitives(&marked_display)
            .iter()
            .any(|p| p["fill"] == "#ff0000")
    );
}

#[test]
fn a_false_hidden_flag_keeps_the_row() {
    let mut kept = three_rows(Some(false));
    drop_hidden_rows(&mut kept);
    assert_eq!(kept.rows.len(), 3);
    let mut unmarked = three_rows(None);
    drop_hidden_rows(&mut unmarked);
    assert_eq!(unmarked.rows.len(), 3);
}

#[test]
fn a_hidden_marked_row_with_visible_content_stays() {
    let paragraph = |text: &str| {
        serde_json::from_value::<LayoutBlock>(json!({
            "kind": "paragraph", "id": "p",
            "runs": [{"kind": "text", "text": text, "fontSize": 10}]
        }))
        .unwrap()
    };
    let mut table = three_rows(Some(true));
    table.rows[1].cells[0].blocks = vec![paragraph("Visible")];
    let before = table.clone();
    drop_hidden_rows(&mut table);
    assert_eq!(table, before);

    table.rows[1].cells[0].blocks = vec![paragraph("")];
    drop_hidden_rows(&mut table);
    assert_eq!(table.rows.len(), 2);
}

#[test]
fn a_fully_hidden_table_has_no_rows_height_or_primitives() {
    let mut hidden = table(vec![row("r0", "#ff0000", 24.0), row("r1", "#ff0000", 4.0)]);
    for row in &mut hidden.rows {
        row.hidden = Some(true);
    }
    drop_hidden_rows(&mut hidden);
    assert!(hidden.rows.is_empty());
    let (measure, display) = render(hidden);
    assert_eq!(measure["totalHeight"], json!(0.0));
    assert!(primitives(&display).is_empty(), "{display}");
}

fn merged_cell(id: &str, span: Option<f64>) -> Value {
    let mut value = cell(id, "#eeeeee");
    if let Some(span) = span {
        value["rowSpan"] = json!(span);
    }
    value
}

fn merge_table(origin_row: usize) -> TableBlock {
    let rows: Vec<Value> = (0..4)
        .map(|index| {
            let mut cells = Vec::new();
            if index == origin_row {
                cells.push(merged_cell("m", Some(3.0)));
            } else if index < origin_row || index > origin_row + 2 {
                cells.push(merged_cell(&format!("a{index}"), None));
            }
            cells.push(merged_cell(&format!("b{index}"), None));
            let mut row = json!({"id": format!("r{index}"), "height": 20, "cells": cells});
            if index == 1 {
                row["hidden"] = json!(true);
            }
            row
        })
        .collect();
    serde_json::from_value(json!({
        "kind": "table", "id": "table", "layoutMode": "fixed",
        "columnWidths": [60, 60], "rows": rows
    }))
    .unwrap()
}

fn ids(table: &TableBlock) -> Vec<Vec<(String, Option<f64>)>> {
    table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| (serde_json::to_string(&cell.id).unwrap(), cell.row_span))
                .collect()
        })
        .collect()
}

#[test]
fn a_merge_through_a_hidden_row_loses_only_that_row() {
    let mut through = merge_table(0);
    drop_hidden_rows(&mut through);
    assert_eq!(
        ids(&through),
        vec![
            vec![("\"m\"".to_owned(), Some(2.0)), ("\"b0\"".to_owned(), None)],
            vec![("\"b2\"".to_owned(), None)],
            vec![("\"a3\"".to_owned(), None), ("\"b3\"".to_owned(), None)],
        ]
    );
    let (measure, display) = render(through);
    assert_eq!(measure["rows"].as_array().unwrap().len(), 3);
    let refs: Vec<(u64, u64, u64)> = primitives(&display)
        .iter()
        .filter(|p| p["cell"].is_object() && p["cell"]["continuation"].is_null())
        .map(|p| {
            let cell = &p["cell"];
            let get = |key: &str| cell[key].as_u64().unwrap();
            (get("row"), get("col"), get("rowSpan"))
        })
        .collect();
    assert!(refs.contains(&(0, 0, 3)), "{refs:?}");
    assert!(refs.contains(&(2, 1, 1)), "{refs:?}");
    assert!(refs.contains(&(3, 0, 1)), "{refs:?}");
    assert!(refs.iter().all(|(row, _, _)| *row != 1), "{refs:?}");
}

#[test]
fn a_merge_starting_in_a_hidden_row_restarts_at_the_next_visible_row() {
    let mut starting = merge_table(1);
    drop_hidden_rows(&mut starting);
    assert_eq!(
        ids(&starting),
        vec![
            vec![("\"a0\"".to_owned(), None), ("\"b0\"".to_owned(), None)],
            vec![("\"m\"".to_owned(), Some(2.0)), ("\"b2\"".to_owned(), None)],
            vec![("\"b3\"".to_owned(), None)],
        ]
    );
    let (measure, _) = render(starting);
    assert_eq!(measure["rows"].as_array().unwrap().len(), 3);
}

#[test]
fn a_nested_table_that_paints_keeps_its_hidden_row() {
    let nested = |cell: Value| {
        serde_json::from_value::<LayoutBlock>(json!({
            "kind": "table", "id": "nested", "columnWidths": [60],
            "rows": [{"id": "n0", "cells": [cell]}]
        }))
        .unwrap()
    };
    let none = json!({"style": "none", "width": 0});
    let empty = json!({"id": "nc", "blocks": [],
        "borders": {"top": none, "right": none, "bottom": none, "left": none}});
    let mut shaded = empty.clone();
    shaded["background"] = json!("#FF0000");
    let mut bordered = empty.clone();
    bordered["borders"]["left"] = json!({"style": "single", "width": 1});
    for (cell, kept) in [(shaded, 3), (bordered, 3), (empty, 2)] {
        let mut table = three_rows(Some(true));
        table.rows[1].cells[0].blocks = vec![nested(cell)];
        drop_hidden_rows(&mut table);
        assert_eq!(table.rows.len(), kept);
    }
}
