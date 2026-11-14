use docx_layout::display_list::{DisplayList, DocAttrs, Primitive, build_display_list_json};
use serde_json::{Value, json};

fn paragraph(id: u64, text: &str) -> Value {
    json!({
        "kind": "paragraph", "id": id, "pmStart": id, "pmEnd": id + 20,
        "runs": [{ "kind": "text", "text": text, "pmStart": id }]
    })
}

fn paragraph_measure(text: &str) -> Value {
    json!({
        "kind": "paragraph", "totalHeight": 24,
        "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": text.len(),
            "width": 80, "ascent": 12, "descent": 4, "lineHeight": 24 }]
    })
}

/// A three by three table; `header` marks the first row as a `w:tblHeader`
/// row, `first_column` sets the table look's first-column flag, and
/// `spanned` makes the first body cell of column 0 span both body rows.
fn table(header: bool, first_column: bool, spanned: bool) -> DisplayList {
    table_with(header, first_column, spanned, false)
}

/// As [`table`]; `continued` marks the first cell of the last row as the
/// continuation of a vertical merge that layout still gives its own slot.
fn table_with(header: bool, first_column: bool, spanned: bool, continued: bool) -> DisplayList {
    let texts = [
        ["Region", "Q1", "Q2"],
        ["North", "12", "19"],
        ["South", "8", "14"],
    ];
    let mut rows = Vec::new();
    let mut measures = Vec::new();
    for (row, cells) in texts.iter().enumerate() {
        let mut block_cells = Vec::new();
        let mut measure_cells = Vec::new();
        for (column, text) in cells.iter().enumerate() {
            if spanned && row == 2 && column == 0 {
                continue;
            }
            let mut cell = json!({
                "blocks": [paragraph(100 * (row as u64 * 3 + column as u64 + 1), text)]
            });
            if spanned && row == 1 && column == 0 {
                cell["rowSpan"] = json!(2);
            }
            if continued && row == 2 && column == 0 {
                cell["mergeContinuation"] = json!(true);
            }
            block_cells.push(cell);
            measure_cells.push(json!({ "blocks": [paragraph_measure(text)] }));
        }
        let mut row_value = json!({ "cells": block_cells });
        if header && row == 0 {
            row_value["isHeader"] = json!(true);
        }
        rows.push(row_value);
        measures.push(json!({ "height": 24, "cells": measure_cells }));
    }
    let mut block = json!({ "kind": "table", "id": 7, "rows": rows });
    if first_column {
        block["firstColumn"] = json!(true);
    }
    let input = json!({
        "measured": [{
            "block": block,
            "measure": { "kind": "table", "columnWidths": [100, 100, 100],
                "totalHeight": 72, "rows": measures }
        }],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400, "h": 300 }, "margins": {}, "fragments": [
                { "kind": "table", "blockId": 7, "x": 50, "y": 50,
                  "width": 300, "height": 72, "rowStart": 0, "rowEnd": 3 }
            ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    serde_json::from_str(&json).unwrap()
}

fn doc_attrs(primitive: &Primitive) -> Option<&DocAttrs> {
    match primitive {
        Primitive::Text(text) => Some(&text.attrs),
        Primitive::GlyphRun(run) => Some(&run.attrs),
        Primitive::Rect(rect) => Some(&rect.attrs),
        Primitive::Image(image) => Some(&image.attrs),
        Primitive::Shape(shape) => Some(&shape.attrs),
        Primitive::Decoration(decoration) => Some(&decoration.attrs),
        Primitive::Line(_) => None,
    }
}

/// `((row, col), heads its column, heads its row, header ids)`.
type CellRow = ((u64, u64), bool, bool, Vec<String>);

/// Per grid cell `(row, col)`: whether it heads its column, whether it heads
/// its row, and the ids of the headers it names, from its text primitive.
fn cells(list: &DisplayList) -> Vec<CellRow> {
    let mut out: Vec<_> = list.pages[0]
        .primitives
        .iter()
        .filter(|primitive| matches!(primitive, Primitive::Text(_)))
        .filter_map(doc_attrs)
        .filter_map(|attrs| attrs.cell.as_deref())
        .map(|cell| {
            (
                (cell.row, cell.col),
                cell.is_header == Some(true),
                cell.is_row_header == Some(true),
                cell.header_ids.clone(),
            )
        })
        .collect();
    out.sort_by_key(|entry| entry.0);
    out.dedup_by_key(|entry| entry.0);
    out
}

fn id(row: u64, col: u64) -> String {
    format!("7-r{row}-c{col}")
}

#[test]
fn header_rows_and_first_column_make_column_zero_head_its_rows() {
    assert_eq!(
        cells(&table(true, true, false)),
        [
            ((0, 0), true, false, vec![]),
            ((0, 1), true, false, vec![]),
            ((0, 2), true, false, vec![]),
            ((1, 0), false, true, vec![]),
            ((1, 1), false, false, vec![id(0, 1), id(1, 0)]),
            ((1, 2), false, false, vec![id(0, 2), id(1, 0)]),
            ((2, 0), false, true, vec![]),
            ((2, 1), false, false, vec![id(0, 1), id(2, 0)]),
            ((2, 2), false, false, vec![id(0, 2), id(2, 0)]),
        ]
    );
}

#[test]
fn first_column_alone_is_formatting_and_marks_no_row_headers() {
    for (row, col) in [(0, 0), (1, 0), (2, 0)] {
        let cell = cells(&table(false, true, false))
            .into_iter()
            .find(|cell| cell.0 == (row, col))
            .unwrap();
        assert_eq!(cell, ((row, col), false, false, vec![]));
    }
}

#[test]
fn header_rows_without_first_column_keep_only_column_headers() {
    let cells = cells(&table(true, false, false));
    assert!(cells.iter().all(|cell| !cell.2));
    assert_eq!(cells[4], ((1, 1), false, false, vec![id(0, 1)]));
}

#[test]
fn a_row_header_spanning_rows_names_every_row_it_covers() {
    let cells = cells(&table(true, true, true));
    assert_eq!(cells[3], ((1, 0), false, true, vec![]));
    assert_eq!(
        cells.iter().find(|cell| cell.0 == (2, 1)).unwrap().3,
        vec![id(0, 1), id(1, 0)]
    );
}

#[test]
fn the_row_header_flag_serializes_camel_case_and_only_when_set() {
    let list = table(true, true, false);
    let json = serde_json::to_value(&list).unwrap().to_string();
    assert!(json.contains("\"isRowHeader\":true"));
    let plain = serde_json::to_value(table(true, false, false))
        .unwrap()
        .to_string();
    assert!(!plain.contains("isRowHeader"));
}

#[test]
fn a_cell_continuing_a_vertical_merge_heads_no_row_even_in_its_own_slot() {
    let cells = cells(&table_with(true, true, false, true));
    let find = |at: (u64, u64)| cells.iter().find(|cell| cell.0 == at).unwrap().clone();
    assert_eq!(find((1, 0)), ((1, 0), false, true, vec![]));
    assert_eq!(find((2, 0)), ((2, 0), false, false, vec![id(0, 0)]));
    assert_eq!(find((2, 1)), ((2, 1), false, false, vec![id(0, 1)]));
}
