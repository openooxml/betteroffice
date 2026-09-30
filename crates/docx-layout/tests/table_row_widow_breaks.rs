//! A row whose cell paragraph widow control or `keepLines` leaves no break in
//! the space left on a page starts on the next one, and one they leave no
//! break in a whole page still splits there, as Word splits such a paragraph.

use serde_json::{Value, json};

fn paragraph(id: usize, lines: usize, attrs: Value) -> (Value, Value) {
    let rows: Vec<_> = (0..lines)
        .map(|_| {
            json!({"headRun":0,"headChar":0,"tailRun":0,"tailChar":1,
                "width":10,"ascent":16,"descent":4,"lineHeight":20})
        })
        .collect();
    (
        json!({"kind":"paragraph","id":id,"runs":[{"kind":"text","text":"x"}],"attrs":attrs}),
        json!({"kind":"paragraph","lines":rows,"totalHeight":20 * lines}),
    )
}

/// A one-cell, one-row table holding a paragraph of `lines` 20px lines.
fn table(lines: usize, attrs: Value) -> Value {
    table_rows(&[(lines, attrs, json!({}))])
}

/// A one-column table whose rows each hold one paragraph of `lines` 20px
/// lines with the given paragraph and row attributes.
fn table_rows(spec: &[(usize, Value, Value)]) -> Value {
    let (mut rows, mut measures) = (Vec::new(), Vec::new());
    for (index, (lines, attrs, row)) in spec.iter().enumerate() {
        let (block, measure) = paragraph(10 + index, *lines, attrs.clone());
        let mut row_block = json!({"id":index,"cells":[{"id":0,"blocks":[block]}]});
        row_block
            .as_object_mut()
            .unwrap()
            .extend(row.as_object().unwrap().clone());
        rows.push(row_block);
        let height = 20 * lines;
        measures.push(
            json!({"height":height,"cells":[{"width":100,"height":height,"blocks":[measure]}]}),
        );
    }
    let height: usize = spec.iter().map(|(lines, ..)| 20 * lines).sum();
    json!({
        "block":{"kind":"table","id":"table","columnWidths":[100],"rows":rows},
        "measure":{"kind":"table","columnWidths":[100],"totalWidth":100,"totalHeight":height,
            "rows":measures}
    })
}

/// Returns table fragments on 100px-tall pages.
fn table_fragments(measured: Vec<Value>, columns: Option<Value>) -> Vec<(usize, Value)> {
    let mut options = json!({"pageSize":{"w":200,"h":120},
        "margins":{"top":10,"right":10,"bottom":10,"left":10}});
    if let Some(columns) = columns {
        options["columns"] = columns;
    }
    let input = json!({"measured":measured,"options":options});
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    layout["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .flat_map(|(page, value)| {
            value["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|fragment| fragment["kind"] == "table")
                .map(move |fragment| (page, fragment.clone()))
                .collect::<Vec<_>>()
        })
        .collect()
}

/// Returns each table fragment's page and height.
fn fragments(measured: Vec<Value>) -> Vec<(usize, f64)> {
    table_fragments(measured, None)
        .into_iter()
        .map(|(page, fragment)| (page, fragment["height"].as_f64().unwrap()))
        .collect()
}

fn shortened_column_fragments(table: Value) -> Vec<(usize, Value)> {
    let (block, measure) = paragraph(1, 4, json!({}));
    table_fragments(
        vec![
            json!({"block":block,"measure":measure}),
            json!({"block":{"kind":"sectionBreak","id":2,"type":"continuous"},
                "measure":{"kind":"sectionBreak"}}),
            table,
        ],
        Some(json!({"count":2,"gap":20})),
    )
}

#[test]
fn a_two_line_widow_controlled_row_skips_shortened_columns() {
    let result = shortened_column_fragments(table(2, json!({"widowControl": true})));
    assert_eq!(result.len(), 1);
    let (page, fragment) = &result[0];
    assert_eq!((*page, fragment["height"].as_f64().unwrap()), (1, 40.0));
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["y"], 10);
    assert!(fragment["clipTop"].is_null());
    assert!(fragment["clipBottom"].is_null());
}

#[test]
fn a_two_line_kept_row_skips_shortened_columns() {
    let result =
        shortened_column_fragments(table(2, json!({"keepLines": true, "widowControl": false})));
    assert_eq!(result.len(), 1);
    let (page, fragment) = &result[0];
    assert_eq!((*page, fragment["height"].as_f64().unwrap()), (1, 40.0));
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["y"], 10);
    assert!(fragment["clipTop"].is_null());
    assert!(fragment["clipBottom"].is_null());
}

#[test]
fn a_tall_row_with_one_line_of_room_starts_on_the_next_page_and_splits_there() {
    let (block, measure) = paragraph(1, 4, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![above, table(10, json!({}))]),
        [(1, 100.0), (2, 100.0)]
    );
}

#[test]
fn a_kept_paragraph_taller_than_a_page_splits_where_it_starts() {
    let (block, measure) = paragraph(1, 2, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![above, table(8, json!({"keepLines": true}))]),
        [(0, 60.0), (1, 100.0)]
    );
}

#[test]
fn a_kept_paragraph_taller_than_a_page_stays_below_its_keep_next_heading() {
    let (block, measure) = paragraph(1, 1, json!({"keepNext": true}));
    let heading = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![
            heading,
            table(8, json!({"keepLines": true, "widowControl": false}))
        ]),
        [(0, 80.0), (1, 80.0)]
    );
}

#[test]
fn a_cant_split_row_taller_than_a_page_still_splits_through_a_kept_paragraph() {
    assert_eq!(
        fragments(vec![table_rows(&[(
            8,
            json!({"keepLines": true}),
            json!({"cantSplit": true})
        )])]),
        [(0, 100.0), (1, 60.0)]
    );
}

#[test]
fn a_cant_split_row_that_fits_a_bare_page_omits_its_repeated_header() {
    assert_eq!(
        fragments(vec![table_rows(&[
            (1, json!({}), json!({"isHeader": true})),
            (
                5,
                json!({"widowControl": false}),
                json!({"cantSplit": true})
            )
        ])]),
        [(0, 20.0), (1, 100.0)]
    );
}

#[test]
fn a_kept_paragraph_taller_than_a_page_splits_below_its_repeated_header() {
    let header = (1, json!({}), json!({"isHeader": true}));
    let kept = (10, json!({"keepLines": true}), json!({}));
    assert_eq!(
        fragments(vec![table_rows(&[header.clone(), kept.clone()])]),
        [(0, 100.0), (1, 100.0), (2, 60.0)]
    );
    let (block, measure) = paragraph(1, 2, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![above, table_rows(&[header, kept])]),
        [(0, 60.0), (1, 100.0), (2, 100.0)]
    );
}

#[test]
fn a_kept_remainder_that_fits_a_bare_page_still_repeats_its_header() {
    assert_eq!(
        fragments(vec![table_rows(&[
            (1, json!({}), json!({"isHeader": true})),
            (9, json!({"keepLines": true}), json!({}))
        ])]),
        [(0, 100.0), (1, 100.0), (2, 40.0)]
    );
}

#[test]
fn an_oversized_cant_split_kept_row_starts_below_its_header_in_the_room_left() {
    let (block, measure) = paragraph(1, 2, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![
            above,
            table_rows(&[
                (1, json!({}), json!({"isHeader": true})),
                (10, json!({"keepLines": true}), json!({"cantSplit": true}))
            ])
        ]),
        [(0, 60.0), (1, 100.0), (2, 100.0)]
    );
}

#[test]
fn a_row_without_widow_control_uses_one_line_of_room_at_the_page_bottom() {
    let (block, measure) = paragraph(1, 4, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![above, table(8, json!({"widowControl": false}))]),
        [(0, 20.0), (1, 100.0), (2, 40.0)]
    );
}

#[test]
fn an_oversized_cant_split_row_without_widow_control_starts_on_a_fresh_page() {
    let (block, measure) = paragraph(1, 4, json!({}));
    let result = table_fragments(
        vec![
            json!({"block":block,"measure":measure}),
            table_rows(&[(
                8,
                json!({"widowControl": false}),
                json!({"cantSplit": true}),
            )]),
        ],
        None,
    );
    assert_eq!(
        result
            .iter()
            .map(|(page, fragment)| (*page, fragment["height"].as_f64().unwrap()))
            .collect::<Vec<_>>(),
        [(1, 100.0), (2, 60.0)]
    );
    for (_, fragment) in &result {
        assert_eq!(fragment["rowStart"], 0);
        assert_eq!(fragment["rowEnd"], 1);
    }
    assert!(result[0].1["clipTop"].is_null());
    assert_eq!(result[0].1["clipBottom"], 100);
    assert_eq!(result[0].1["clipBottom"], result[1].1["clipTop"]);
    assert!(result[1].1["clipBottom"].is_null());
}

#[test]
fn an_oversized_cant_split_row_keeps_widow_control_where_a_fresh_page_allows_it() {
    let (block, measure) = paragraph(1, 4, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![
            above,
            table_rows(&[(8, json!({}), json!({"cantSplit": true}))])
        ]),
        [(1, 100.0), (2, 60.0)]
    );
}

/// Pages of the keepNext heading and of the table's first fragment after a
/// 60px filler on 100px pages.
fn heading_and_table_pages(table: Value) -> (Option<usize>, Option<usize>) {
    heading_and_table_pages_after(3, table)
}

fn heading_and_table_pages_after(
    filler_lines: usize,
    table: Value,
) -> (Option<usize>, Option<usize>) {
    let (filler, filler_measure) = paragraph(1, filler_lines, json!({}));
    let (heading, heading_measure) = paragraph(2, 1, json!({"keepNext": true}));
    let input = json!({
        "measured":[
            {"block":filler,"measure":filler_measure},
            {"block":heading,"measure":heading_measure},
            table,
        ],
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10}}
    });
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let page_of = |matches: &dyn Fn(&Value) -> bool| {
        layout["pages"]
            .as_array()
            .unwrap()
            .iter()
            .position(|page| page["fragments"].as_array().unwrap().iter().any(matches))
    };
    (
        page_of(&|fragment| fragment["blockId"] == 2),
        page_of(&|fragment| fragment["kind"] == "table"),
    )
}

fn header_table_over_a_kept_row() -> Value {
    table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (
            8,
            json!({"keepLines": true, "widowControl": false}),
            json!({}),
        ),
    ])
}

#[test]
fn a_keep_next_heading_moves_with_a_header_table_over_a_kept_row() {
    assert_eq!(
        heading_and_table_pages(header_table_over_a_kept_row()),
        (Some(1), Some(1))
    );
}

#[test]
fn a_keep_next_heading_moves_with_an_oversized_widow_controlled_row() {
    assert_eq!(
        heading_and_table_pages(table(8, json!({"widowControl": true}))),
        (Some(1), Some(1))
    );
}

#[test]
fn a_keep_next_heading_moves_with_a_tall_floating_table_placed_in_flow() {
    let mut table = header_table_over_a_kept_row();
    table["block"]["floating"] =
        json!({"horzAnchor": "margin", "vertAnchor": "text", "tblpX": 0, "tblpY": 0});
    assert_eq!(heading_and_table_pages(table), (Some(1), Some(1)));
}

#[test]
fn a_keep_next_heading_stays_beside_a_fitting_floating_table() {
    let mut table = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (2, json!({"widowControl": true}), json!({})),
    ]);
    table["block"]["floating"] =
        json!({"horzAnchor": "page", "tblpX": 90, "vertAnchor": "page", "tblpY": 10});
    assert_eq!(heading_and_table_pages_after(2, table).0, Some(0));
}

#[test]
fn a_keep_next_heading_stays_beside_a_text_floating_table_lifted_above_it() {
    let mut table = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (2, json!({"widowControl": true}), json!({})),
    ]);
    table["block"]["floating"] =
        json!({"horzAnchor": "page", "tblpX": 90, "vertAnchor": "text", "tblpY": -20});
    assert_eq!(heading_and_table_pages_after(2, table), (Some(0), Some(0)));
}

#[test]
fn a_full_width_text_floating_table_splits_its_rows_at_any_line() {
    let mut table = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (2, json!({"widowControl": true}), json!({})),
    ]);
    table["block"]["columnWidths"] = json!([180]);
    table["measure"]["columnWidths"] = json!([180]);
    table["measure"]["totalWidth"] = json!(180);
    for row in table["measure"]["rows"].as_array_mut().unwrap() {
        row["cells"][0]["width"] = json!(180);
    }
    table["block"]["floating"] =
        json!({"horzAnchor": "margin", "vertAnchor": "text", "tblpX": 0, "tblpY": 0});
    assert_eq!(heading_and_table_pages_after(2, table), (Some(0), Some(0)));
}

#[test]
fn a_kept_row_follows_its_next_row_to_a_taller_page() {
    let mut table = table_rows(&[
        (1, json!({}), json!({})),
        (10, json!({}), json!({"height": 200, "heightRule": "exact"})),
        (
            2,
            json!({"keepNext": true, "widowControl": false}),
            json!({}),
        ),
        (
            10,
            json!({"keepLines": true, "widowControl": false}),
            json!({}),
        ),
    ]);
    table["block"]["rows"][1]["cells"][0]["blocks"] = json!([]);
    table["measure"]["rows"][1]["cells"][0]["blocks"] = json!([]);
    let input = json!({"measured": [table], "options": {
        "pageSize": {"w": 200, "h": 320},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
        "sectionPageFloatBands": [{"default": [], "first": [{"top": 190, "bottom": 230}]}],
    }});
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let page_of = |row: u64| {
        layout["pages"]
            .as_array()
            .unwrap()
            .iter()
            .position(|page| {
                page["fragments"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|fragment| {
                        fragment["kind"] == "table"
                            && fragment["rowStart"].as_u64().unwrap() <= row
                            && row < fragment["rowEnd"].as_u64().unwrap()
                    })
            })
            .unwrap()
    };
    assert_eq!(page_of(2), page_of(3));
}

#[test]
fn a_kept_row_stays_when_the_next_row_cannot_split_on_any_page() {
    let (block, measure) = paragraph(1, 3, json!({}));
    let table = table_rows(&[
        (
            2,
            json!({"keepNext": true, "widowControl": false}),
            json!({}),
        ),
        (
            8,
            json!({"keepLines": true, "widowControl": false}),
            json!({}),
        ),
    ]);
    let fragments = fragments(vec![json!({"block": block, "measure": measure}), table]);
    assert_eq!(fragments[0].0, 0);
}
