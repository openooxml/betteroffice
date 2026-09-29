//! A keep-with-next run moves to the next page only when it would fit there
//! whole, measured as placement lays it out on a fresh page.

use serde_json::{Value, json};

fn paragraph(id: u32, lines: &[f64], attrs: Value) -> Value {
    let before = attrs["spacing"]["before"].as_f64().unwrap_or(0.0);
    let after = attrs["spacing"]["after"].as_f64().unwrap_or(0.0);
    let runs = if lines.is_empty() {
        json!([])
    } else {
        json!([{"kind": "text", "text": "x", "fmt": {}}])
    };
    let lines: Vec<_> = lines
        .iter()
        .map(|height| {
            json!({"headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1, "width": 10,
                "ascent": 8, "descent": 2, "lineHeight": height})
        })
        .collect();
    json!({
        "block": {"kind": "paragraph", "id": id, "runs": runs, "attrs": attrs},
        "measure": {"kind": "paragraph", "lines": lines,
            "totalHeight": lines.iter().map(|line| line["lineHeight"].as_f64().unwrap()).sum::<f64>() + before + after}
    })
}

/// Lays `measured` out on 100px-tall pages and returns whether blocks 2 and 3
/// share a page.
fn head_stays_with_follower(measured: Vec<Value>) -> bool {
    let input = json!({"measured": measured, "options": {
        "pageSize": {"w": 200, "h": 120},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}}});
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_json(&input.to_string()).unwrap()).unwrap();
    let page_of = |id: f64| {
        layout["pages"].as_array().unwrap().iter().position(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .any(|fragment| fragment["blockId"].as_f64() == Some(id))
        })
    };
    page_of(2.0) == page_of(3.0)
}

#[test]
fn spacing_owed_at_the_cursor_does_not_count_on_a_fresh_page() {
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[30.0], json!({"spacing": {"after": 25}})),
        paragraph(2, &[20.0], json!({"keepNext": true})),
        paragraph(3, &[20.0, 20.0, 20.0], json!({"keepLines": true})),
    ]));
}

#[test]
fn an_empty_follower_still_needs_the_gap_above_it() {
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[70.0], json!({})),
        paragraph(
            2,
            &[20.0],
            json!({"keepNext": true, "spacing": {"after": 15}})
        ),
        paragraph(3, &[], json!({})),
    ]));
}

#[test]
fn a_table_follower_needs_its_kept_leading_rows() {
    let cells: Vec<Value> = (0..2)
        .map(|index| paragraph(10 + index, &[20.0], json!({"keepNext": index == 0})))
        .collect();
    let rows: Vec<_> = (0..2)
        .map(|index| {
            json!({"id": 20 + index, "cantSplit": true,
                "cells": [{"id": 30 + index, "blocks": [cells[index]["block"].clone()]}]})
        })
        .collect();
    let measured_rows: Vec<_> = (0..2)
        .map(|index| {
            json!({"height": 20, "cells": [{"width": 100, "height": 20,
                "blocks": [cells[index]["measure"].clone()]}]})
        })
        .collect();
    let table = json!({
        "block": {"kind": "table", "id": 3, "columnWidths": [100], "rows": rows},
        "measure": {"kind": "table", "totalWidth": 100, "totalHeight": 40,
            "columnWidths": [100], "rows": measured_rows}
    });
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[55.0], json!({})),
        paragraph(
            2,
            &[10.0],
            json!({"keepNext": true, "spacing": {"after": 10}})
        ),
        table,
    ]));
}
