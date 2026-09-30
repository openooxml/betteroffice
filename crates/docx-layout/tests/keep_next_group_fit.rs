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

/// Lays `measured` out on 100px-tall pages and returns each block's first
/// page and column.
fn place(measured: Vec<Value>, columns: Option<Value>) -> impl Fn(f64) -> Option<(usize, f64)> {
    let mut options = json!({"pageSize": {"w": 200, "h": 120},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}});
    if let Some(columns) = columns {
        options["columns"] = columns;
    }
    let input = json!({"measured": measured, "options": options});
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_json(&input.to_string()).unwrap()).unwrap();
    move |id: f64| {
        layout["pages"]
            .as_array()
            .unwrap()
            .iter()
            .enumerate()
            .find_map(|(index, page)| {
                page["fragments"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|fragment| fragment["blockId"].as_f64() == Some(id))
                    .map(|fragment| (index, fragment["x"].as_f64().unwrap()))
            })
    }
}

/// Whether blocks 2 and 3 share a page.
fn head_stays_with_follower(measured: Vec<Value>) -> bool {
    let at = place(measured, None);
    at(2.0).map(|(page, _)| page) == at(3.0).map(|(page, _)| page)
}

/// A table of 20px single-cell rows; `rows` gives each row's header flag,
/// whether its paragraph keeps with the next row, and whether it may split.
fn table(id: u32, rows: &[(bool, bool, bool)], floating: Option<Value>) -> Value {
    let cells: Vec<Value> = rows
        .iter()
        .enumerate()
        .map(|(index, (_, keep_next, _))| {
            paragraph(100 + index as u32, &[20.0], json!({"keepNext": keep_next}))
        })
        .collect();
    let block_rows: Vec<_> = rows
        .iter()
        .enumerate()
        .map(|(index, (header, _, cant_split))| {
            json!({"id": 200 + index, "isHeader": header, "cantSplit": cant_split,
                "cells": [{"id": 300 + index, "blocks": [cells[index]["block"].clone()]}]})
        })
        .collect();
    let measured_rows: Vec<_> = cells
        .iter()
        .map(|cell| {
            json!({"height": 20, "cells": [{"width": 100, "height": 20,
                "blocks": [cell["measure"].clone()]}]})
        })
        .collect();
    let mut block = json!({"kind": "table", "id": id, "columnWidths": [100], "rows": block_rows});
    if let Some(floating) = floating {
        block["floating"] = floating;
    }
    json!({
        "block": block,
        "measure": {"kind": "table", "totalWidth": 100, "totalHeight": 20.0 * rows.len() as f64,
            "columnWidths": [100], "rows": measured_rows}
    })
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
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[55.0], json!({})),
        paragraph(
            2,
            &[10.0],
            json!({"keepNext": true, "spacing": {"after": 10}})
        ),
        table(3, &[(false, true, true), (false, false, true)], None),
    ]));
}

#[test]
fn a_row_chain_taller_than_a_page_does_not_hold_the_heading_back() {
    let mut rows = table(3, &[(false, true, true), (false, false, true)], None);
    rows["measure"]["rows"][1]["height"] = json!(90);
    rows["measure"]["totalHeight"] = json!(110);
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[75.0], json!({})),
        paragraph(2, &[10.0], json!({"keepNext": true})),
        rows,
    ]));
}

#[test]
fn a_row_chain_starting_in_the_header_band_comes_along() {
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[25.0], json!({})),
        paragraph(2, &[10.0], json!({"keepNext": true})),
        table(
            3,
            &[
                (true, false, true),
                (true, true, true),
                (false, true, true),
                (false, false, true)
            ],
            None
        ),
    ]));
}

#[test]
fn a_floating_table_follower_is_not_weighed_as_flow() {
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[60.0], json!({})),
        paragraph(2, &[10.0], json!({"keepNext": true})),
        table(
            3,
            &[(false, true, true), (false, false, true)],
            Some(json!({"horzAnchor": "page", "tblpX": 90, "vertAnchor": "page", "tblpY": 10}))
        ),
    ]));
}

#[test]
fn a_run_that_does_not_fit_moves_to_the_next_column() {
    let at = place(
        vec![
            paragraph(1, &[30.0], json!({"spacing": {"after": 80}})),
            paragraph(2, &[20.0], json!({"keepNext": true})),
            paragraph(3, &[10.0], json!({})),
            json!({"block": {"kind": "columnBreak", "id": 4}, "measure": {"kind": "columnBreak"}}),
            paragraph(5, &[10.0], json!({})),
        ],
        Some(json!({"count": 2, "gap": 20})),
    );
    let (heading, follower) = (at(2.0).unwrap(), at(3.0).unwrap());
    assert_eq!(heading, follower);
    assert_eq!(heading.0, 0);
    assert!(heading.1 > at(1.0).unwrap().1);
}

#[test]
fn a_row_chain_that_only_fits_without_the_header_rows_is_not_weighed() {
    let mut rows = table(
        3,
        &[
            (true, false, true),
            (true, false, true),
            (false, true, true),
            (false, false, true),
        ],
        None,
    );
    rows["measure"]["rows"][3]["height"] = json!(60);
    rows["measure"]["totalHeight"] = json!(120);
    assert!(head_stays_with_follower(vec![
        paragraph(1, &[40.0], json!({})),
        paragraph(2, &[10.0], json!({"keepNext": true})),
        rows,
    ]));
}

#[test]
fn a_run_skips_columns_too_short_for_it() {
    let at = place(
        vec![
            paragraph(1, &[60.0], json!({})),
            json!({"block": {"kind": "sectionBreak", "id": 4, "type": "continuous"},
                "measure": {"kind": "sectionBreak"}}),
            paragraph(2, &[10.0], json!({"keepNext": true})),
            paragraph(3, &[40.0], json!({})),
        ],
        Some(json!({"count": 2, "gap": 20})),
    );
    assert_eq!(at(2.0).unwrap().0, 1);
    assert_eq!(at(3.0).unwrap().0, 1);
}
