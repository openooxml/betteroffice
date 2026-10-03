//! A row whose cell paragraph widow control or `keepLines` leaves no break in
//! the space left on a page starts on the next one, and one they leave no
//! break in a whole page still splits there, as Word splits such a paragraph.

use serde_json::{Value, json};

#[path = "fixtures/nested_table_cell_window.rs"]
mod nested_table_cell_window;

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

fn two_cell_table(spec: &[(usize, f64, Value)]) -> Value {
    let (mut cells, mut measures) = (Vec::new(), Vec::new());
    let mut row_height = 0.0_f64;
    for (cell, (lines, before, attrs)) in spec.iter().enumerate() {
        let (mut block, measure) = paragraph(10 + cell, *lines, attrs.clone());
        block["attrs"]["spacing"] = json!({"before": before});
        let height = 20.0 * *lines as f64 + before;
        row_height = row_height.max(height);
        cells.push(json!({"id": cell, "blocks": [block]}));
        measures.push(json!({"width": 100, "height": height, "blocks": [measure]}));
    }
    json!({
        "block": {"kind": "table", "id": "table", "columnWidths": [100, 100],
            "rows": [{"id": 0, "cells": cells}]},
        "measure": {"kind": "table", "columnWidths": [100, 100], "totalWidth": 200,
            "totalHeight": row_height, "rows": [{"height": row_height, "cells": measures}]}
    })
}

fn after_filler(lines: usize, table: Value) -> Vec<Value> {
    let (block, measure) = paragraph(1, lines, json!({}));
    vec![json!({"block": block, "measure": measure}), table]
}

/// (page, rowStart, rowEnd) of each table fragment of `table` after a five-line
/// filler on 200 px tall pages; `legacy` also asserts that none carries cellClips.
fn header_table_fragments(table: Value, legacy: bool) -> Vec<(usize, u64, u64)> {
    let input = json!({
        "measured": after_filler(5, table),
        "options": {"pageSize": {"w": 240, "h": 220},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}},
    });
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let mut fragments = Vec::new();
    for (page, value) in layout["pages"].as_array().unwrap().iter().enumerate() {
        for fragment in value["fragments"].as_array().unwrap() {
            if fragment["kind"] == "table" {
                assert!(!legacy || fragment.get("cellClips").is_none());
                let row = |key: &str| fragment[key].as_u64().unwrap();
                fragments.push((page, row("rowStart"), row("rowEnd")));
            }
        }
    }
    fragments
}

fn cell_windows(fragment: &Value) -> Vec<(f64, f64)> {
    fragment["cellClips"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .map(|(cell, clip)| {
            assert_eq!(clip["cell"], cell);
            (
                clip["top"].as_f64().unwrap(),
                clip["bottom"].as_f64().unwrap(),
            )
        })
        .collect()
}

fn placed_cell_lines(measured: &[Value]) -> Vec<Vec<(usize, std::ops::Range<usize>)>> {
    use docx_layout::placement::{PlacedItem, PlacementInput, place_layout};
    let mut measured = measured.to_vec();
    for (index, entry) in measured.iter_mut().enumerate() {
        entry["block"]["pmStart"] = json!(2 * index);
        entry["block"]["pmEnd"] = json!(2 * index + 1);
    }
    let mut input: docx_layout::types::Input = serde_json::from_value(json!({
        "measured": measured, "options": {"pageSize": {"w": 200, "h": 120},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}}
    }))
    .unwrap();
    let layout = docx_layout::compute_layout_input(&mut input).unwrap();
    let placed = place_layout(&PlacementInput {
        layout: &layout,
        measured: &input.measured,
        headers_footers: None,
        bands_composed: false,
        notes: &[],
    });
    assert!(placed.issues.is_empty(), "{:?}", placed.issues);
    placed
        .pages
        .iter()
        .map(|page| {
            page.regions
                .iter()
                .flat_map(|region| &region.items)
                .filter_map(|item| match item {
                    PlacedItem::Paragraph(paragraph) => {
                        let id = serde_json::to_value(&paragraph.block.id)
                            .unwrap()
                            .as_f64()? as usize;
                        id.checked_sub(10)
                            .map(|cell| (cell, paragraph.lines.clone()))
                    }
                    _ => None,
                })
                .collect()
        })
        .collect()
}

#[test]
fn offset_cell_line_grids_split_and_resume_at_independent_boundaries() {
    let measured = after_filler(
        2,
        two_cell_table(&[(6, 0.0, json!({})), (6, 4.4, json!({}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(
        result.iter().map(|(page, _)| *page).collect::<Vec<_>>(),
        [0, 1]
    );
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 44.4)]);
    assert_eq!(cell_windows(&result[1].1), [(60.0, 120.0), (44.4, 124.4)]);
    assert_eq!(result[0].1["height"], 60);
    assert_eq!(result[1].1["height"], 80);
    let lines = placed_cell_lines(&measured);
    assert_eq!(
        lines,
        [vec![(0, 0..3), (1, 0..2)], vec![(0, 3..6), (1, 2..6)]]
    );
    for cell in 0..2 {
        let seen: Vec<_> = lines
            .iter()
            .flatten()
            .filter(|(index, _)| *index == cell)
            .flat_map(|(_, lines)| lines.clone())
            .collect();
        assert_eq!(seen, (0..6).collect::<Vec<_>>());
    }
}

#[test]
fn offset_cell_line_grids_move_the_whole_row_below_a_float_band_on_the_same_page() {
    let input = json!({
        "measured": after_filler(
            6,
            two_cell_table(&[(6, 0.0, json!({})), (6, 4.4, json!({}))]),
        ),
        "options": {
            "pageSize": {"w": 240, "h": 500},
            "margins": {"top": 20, "right": 10, "bottom": 10, "left": 10},
            "sectionPageFloatBands": [{"default": [], "first": [{"top": 200, "bottom": 250}]}],
        },
    });
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let pages = layout["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments: Vec<_> = pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(fragments.len(), 1);
    let fragment = fragments[0];
    assert_eq!(fragment["y"], 250);
    assert!((fragment["height"].as_f64().unwrap() - 124.4).abs() < 0.01);
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert!(fragment["clipTop"].is_null());
    assert!(fragment["clipBottom"].is_null());
    assert!(fragment["cellClips"].is_null());
}

fn offset_cell_float_band_fragments(leading_row: Value) -> Vec<Value> {
    let mut table = two_cell_table(&[(6, 0.0, json!({})), (6, 4.4, json!({}))]);
    table["block"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, leading_row["block"]["rows"][0].clone());
    table["measure"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, leading_row["measure"]["rows"][0].clone());
    table["measure"]["totalHeight"] = json!(144.4);
    let input = json!({
        "measured": after_filler(6, table),
        "options": {
            "pageSize": {"w": 240, "h": 500},
            "margins": {"top": 20, "right": 10, "bottom": 10, "left": 10},
            "sectionPageFloatBands": [{"default": [], "first": [{"top": 200, "bottom": 250}]}],
        },
    });
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let pages = layout["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .cloned()
        .collect()
}

#[test]
fn repeated_headers_move_with_offset_cells_below_a_float_band_on_the_same_page() {
    let header = table_rows(&[(1, json!({}), json!({"isHeader": true}))]);
    let fragments = offset_cell_float_band_fragments(header);
    assert_eq!(fragments.len(), 1);
    let fragment = &fragments[0];
    assert_eq!(fragment["y"], 250);
    assert!((fragment["height"].as_f64().unwrap() - 144.4).abs() < 0.01);
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 2);
    assert!(fragment["headerRowCount"].is_null());
    assert!(fragment["clipTop"].is_null());
    assert!(fragment["clipBottom"].is_null());
    assert!(fragment["cellClips"].is_null());
}

#[test]
fn keep_next_rows_move_with_offset_cells_below_a_float_band_on_the_same_page() {
    let kept = table_rows(&[(1, json!({"keepNext": true}), json!({}))]);
    let fragments = offset_cell_float_band_fragments(kept);
    assert_eq!(fragments.len(), 1);
    let fragment = &fragments[0];
    assert_eq!(fragment["y"], 250);
    assert!((fragment["height"].as_f64().unwrap() - 144.4).abs() < 0.01);
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 2);
    assert!(fragment["clipTop"].is_null());
    assert!(fragment["clipBottom"].is_null());
    assert!(fragment["cellClips"].is_null());
}

#[test]
fn a_keep_next_heading_moves_with_offset_cells_and_headers_below_a_float_band() {
    let header = table_rows(&[(1, json!({}), json!({"isHeader": true}))]);
    let mut table = two_cell_table(&[(6, 0.0, json!({})), (6, 4.4, json!({}))]);
    table["block"]["rows"] = json!([header["block"]["rows"][0], table["block"]["rows"][0],]);
    table["measure"]["rows"] = json!([header["measure"]["rows"][0], table["measure"]["rows"][0],]);
    table["measure"]["totalHeight"] = json!(144.4);
    let (block, measure) = paragraph(2, 1, json!({"keepNext": true}));
    let mut measured = after_filler(5, json!({"block": block, "measure": measure}));
    measured.push(table);
    let input = json!({
        "measured": measured,
        "options": {
            "pageSize": {"w": 240, "h": 500},
            "margins": {"top": 20, "right": 10, "bottom": 10, "left": 10},
            "sectionPageFloatBands": [{"default": [], "first": [{"top": 200, "bottom": 250}]}],
        },
    });
    let layout: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    let pages = layout["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments = pages[0]["fragments"].as_array().unwrap();
    let heading = fragments
        .iter()
        .find(|fragment| fragment["blockId"] == 2)
        .unwrap();
    assert_eq!(heading["y"], 250);
    let tables: Vec<_> = fragments
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0]["y"], 270);
    assert_eq!(tables[0]["rowStart"], 0);
    assert_eq!(tables[0]["rowEnd"], 2);
}

#[test]
fn a_four_line_cell_splits_two_and_two_while_the_other_cell_finishes() {
    let measured = after_filler(
        2,
        two_cell_table(&[(4, 0.0, json!({})), (2, 4.4, json!({}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 40.0), (0.0, 44.4)]);
    assert_eq!(cell_windows(&result[1].1), [(40.0, 80.0), (44.4, 44.4)]);
    assert_eq!(
        placed_cell_lines(&measured),
        [vec![(0, 0..2), (1, 0..2)], vec![(0, 2..4)]]
    );
}

#[test]
fn a_three_line_widow_controlled_cell_stays_whole_while_another_splits() {
    let measured = after_filler(
        2,
        two_cell_table(&[(4, 0.0, json!({})), (3, 4.4, json!({}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(result[0].0, 0);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 40.0), (0.0, 0.0)]);
    assert_eq!(cell_windows(&result[1].1), [(40.0, 80.0), (0.0, 64.4)]);
    assert_eq!(
        placed_cell_lines(&measured),
        [vec![(0, 0..2)], vec![(0, 2..4), (1, 0..3)]]
    );
}

#[test]
fn per_cell_continuations_keep_each_cursor_across_three_pages() {
    let measured = after_filler(
        2,
        two_cell_table(&[(10, 0.0, json!({})), (10, 4.4, json!({}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(result.len(), 3);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 44.4)]);
    assert_eq!(cell_windows(&result[1].1), [(60.0, 160.0), (44.4, 144.4)]);
    assert_eq!(cell_windows(&result[2].1), [(160.0, 200.0), (144.4, 204.4)]);
    assert_eq!(
        placed_cell_lines(&measured),
        [
            vec![(0, 0..3), (1, 0..2)],
            vec![(0, 3..8), (1, 2..7)],
            vec![(0, 8..10), (1, 7..10)],
        ]
    );
}

#[test]
fn matching_cell_boundaries_keep_the_exact_legacy_fragments() {
    let one = table_fragments(after_filler(2, table(6, json!({}))), None);
    let two = table_fragments(
        after_filler(
            2,
            two_cell_table(&[(6, 0.0, json!({})), (6, 0.0, json!({}))]),
        ),
        None,
    );
    for ((one_page, one), (two_page, two)) in one.iter().zip(&two) {
        assert_eq!(one_page, two_page);
        let mut expected = one.clone();
        expected["width"] = json!(200);
        assert_eq!(
            serde_json::to_string(two).unwrap(),
            serde_json::to_string(&expected).unwrap()
        );
        assert!(two.get("cellClips").is_none());
    }
    assert_eq!(two.len(), 2);
    assert_eq!(two[0].1["clipBottom"], 60);
    assert_eq!(two[1].1["clipTop"], 60);
    assert!(two[1].1.get("clipBottom").is_none());
}

#[test]
fn unsupported_cell_windows_preserve_the_legacy_row_path() {
    let original = two_cell_table(&[(4, 0.0, json!({})), (4, 4.4, json!({}))]);
    for exclusion in [
        "cantSplit",
        "exact",
        "minimum",
        "header",
        "span",
        "center",
        "bottom",
        "rotated",
        "footnote",
        "endnote",
        "floatingImage",
        "wrappedImage",
        "behindImage",
        "floatImage",
        "anchoredImage",
        "nestedFloatingImage",
        "floatingTable",
        "anchoredShape",
        "unanchoredShape",
        "anchoredTextBox",
        "anchoredChart",
        "anchoredImageBlock",
    ] {
        let mut table = original.clone();
        match exclusion {
            "cantSplit" => table["block"]["rows"][0]["cantSplit"] = json!(true),
            "exact" => {
                table["block"]["rows"][0]["height"] = json!(84.4);
                table["block"]["rows"][0]["heightRule"] = json!("exact");
            }
            "minimum" => {
                table["block"]["rows"][0]["height"] = json!(90);
                table["measure"]["rows"][0]["height"] = json!(90);
                table["measure"]["totalHeight"] = json!(90);
            }
            "header" => table["block"]["rows"][0]["isHeader"] = json!(true),
            "span" => table["block"]["rows"][0]["cells"][0]["rowSpan"] = json!(2),
            "center" | "bottom" => {
                table["block"]["rows"][0]["cells"][0]["verticalAlign"] = json!(exclusion)
            }
            "rotated" => table["block"]["rows"][0]["cells"][0]["textDirection"] = json!("tbRl"),
            "footnote" | "endnote" => {
                let key = if exclusion == "footnote" {
                    "footnoteRefId"
                } else {
                    "endnoteRefId"
                };
                table["block"]["rows"][0]["cells"][0]["blocks"][0]["runs"][0][key] = json!(1);
            }
            "floatingImage"
            | "wrappedImage"
            | "behindImage"
            | "floatImage"
            | "anchoredImage"
            | "nestedFloatingImage"
            | "floatingTable" => {
                let mut image = json!({"kind": "image", "src": "synthetic-image",
                    "width": 20, "height": 20, "wrapType": "inFront",
                    "position": {"vertical": {"relativeTo": "paragraph", "posOffset": 50}}});
                match exclusion {
                    "wrappedImage" | "behindImage" => {
                        image["wrapType"] = json!(if exclusion == "behindImage" {
                            "behind"
                        } else {
                            "square"
                        });
                        image.as_object_mut().unwrap().remove("position");
                    }
                    "floatImage" => {
                        image.as_object_mut().unwrap().remove("wrapType");
                        image.as_object_mut().unwrap().remove("position");
                        image["displayMode"] = json!("float");
                    }
                    "anchoredImage" => {
                        image.as_object_mut().unwrap().remove("wrapType");
                    }
                    _ => {}
                }
                if matches!(exclusion, "nestedFloatingImage" | "floatingTable") {
                    let mut nested = table_rows(&[(4, json!({}), json!({}))]);
                    nested["block"]["id"] = json!("nested");
                    if exclusion == "floatingTable" {
                        nested["block"]["floating"] = json!({
                            "vertAnchor": "text", "horzAnchor": "text", "tblpY": 0
                        });
                    } else {
                        nested["block"]["rows"][0]["cells"][0]["blocks"][0]["runs"]
                            .as_array_mut()
                            .unwrap()
                            .push(image);
                    }
                    table["block"]["rows"][0]["cells"][0]["blocks"] = json!([nested["block"]]);
                    table["measure"]["rows"][0]["cells"][0]["blocks"] = json!([nested["measure"]]);
                } else {
                    table["block"]["rows"][0]["cells"][0]["blocks"][0]["runs"]
                        .as_array_mut()
                        .unwrap()
                        .push(image);
                }
            }
            "anchoredShape" | "unanchoredShape" | "anchoredTextBox" | "anchoredChart"
            | "anchoredImageBlock" => {
                let kind = match exclusion {
                    "anchoredShape" | "unanchoredShape" => "shape",
                    "anchoredTextBox" => "textBox",
                    "anchoredChart" => "chart",
                    _ => "image",
                };
                let mut drawing = json!({"kind": kind, "id": 50, "width": 20, "height": 10,
                    "position": {"vertical": {"relativeTo": "paragraph", "posOffset": 50}}});
                match exclusion {
                    "anchoredShape" | "unanchoredShape" => {
                        drawing["shapeType"] = json!("rect");
                        drawing["geometryPath"] = json!([]);
                        drawing["children"] = json!([]);
                        if exclusion == "anchoredShape" {
                            drawing["wrapType"] = json!("behind");
                        } else {
                            drawing.as_object_mut().unwrap().remove("position");
                        }
                    }
                    "anchoredTextBox" => drawing["content"] = json!([]),
                    "anchoredChart" => drawing["chart"] = json!({}),
                    _ => {
                        drawing["src"] = json!("synthetic-image");
                        drawing["anchor"] = json!({"isAnchored": true});
                    }
                }
                table["block"]["rows"][0]["cells"][0]["blocks"]
                    .as_array_mut()
                    .unwrap()
                    .push(drawing);
                table["measure"]["rows"][0]["cells"][0]["blocks"]
                    .as_array_mut()
                    .unwrap()
                    .push(json!({"kind": kind, "width": 20, "height": 10, "innerMeasures": []}));
                if exclusion != "anchoredShape" {
                    table["measure"]["rows"][0]["cells"][0]["height"] = json!(90);
                    table["measure"]["rows"][0]["height"] = json!(90);
                    table["measure"]["totalHeight"] = json!(90);
                }
            }
            _ => unreachable!(),
        }
        let height = table["measure"]["totalHeight"].clone();
        let result = table_fragments(after_filler(2, table), None);
        assert_eq!(result.len(), 1, "{exclusion}");
        assert!(
            result
                .iter()
                .all(|(_, fragment)| fragment.get("cellClips").is_none()),
            "{exclusion}"
        );
        assert_eq!(result[0].0, 1, "{exclusion}");
        assert_eq!(result[0].1["height"], height, "{exclusion}");
    }
}

#[test]
fn a_nested_table_stays_atomic_when_another_cell_splits() {
    let input = nested_table_cell_window::input();
    let measured = input["measured"].as_array().unwrap().clone();
    let result = table_fragments(measured, None);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 20.0)]);
    assert_eq!(cell_windows(&result[1].1), [(60.0, 120.0), (20.0, 80.0)]);
}

#[test]
fn repeated_headers_reserve_their_band_before_per_cell_cuts() {
    let header = table_rows(&[(1, json!({}), json!({"isHeader": true}))]);
    let mut table = two_cell_table(&[(6, 0.0, json!({})), (6, 4.4, json!({}))]);
    table["block"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, header["block"]["rows"][0].clone());
    table["measure"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, header["measure"]["rows"][0].clone());
    table["measure"]["totalHeight"] = json!(144.4);
    let result = table_fragments(after_filler(1, table), None);
    assert_eq!(result.len(), 2);
    assert_eq!(result[0].0, 0);
    assert_eq!(result[0].1["height"], 80);
    assert_eq!(result[1].1["height"], 100);
    assert_eq!(result[1].1["headerRowCount"], 1);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 44.4)]);
    assert_eq!(cell_windows(&result[1].1), [(60.0, 120.0), (44.4, 124.4)]);
    assert!(
        result
            .iter()
            .all(|(_, fragment)| fragment["cellClips"][0]["row"] == 1)
    );
}

#[test]
fn repeated_headers_move_with_an_offset_cell_keep_next_chain_without_a_header_only_fragment() {
    let ends = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (1, json!({}), json!({})),
    ]);
    let mut table = two_cell_table(&[
        (6, 0.0, json!({"keepNext": true})),
        (6, 4.4, json!({"keepNext": true})),
    ]);
    table["block"]["rows"] = json!([
        ends["block"]["rows"][0],
        table["block"]["rows"][0],
        ends["block"]["rows"][1],
    ]);
    table["measure"]["rows"] = json!([
        ends["measure"]["rows"][0],
        table["measure"]["rows"][0],
        ends["measure"]["rows"][1],
    ]);
    table["measure"]["totalHeight"] = json!(164.4);
    assert_eq!(header_table_fragments(table, false), [(1, 0, 3)]);
}

#[test]
fn repeated_headers_move_with_a_keep_next_chain_starting_in_the_header_band() {
    let ends = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (1, json!({"keepNext": true}), json!({"isHeader": true})),
        (1, json!({}), json!({})),
    ]);
    let keep = json!({"keepNext": true});
    let mut table = two_cell_table(&[(6, 0.0, keep.clone()), (6, 4.4, keep)]);
    for key in ["block", "measure"] {
        table[key]["rows"] = json!([
            ends[key]["rows"][0],
            ends[key]["rows"][1],
            table[key]["rows"][0],
            ends[key]["rows"][2],
        ]);
    }
    table["measure"]["totalHeight"] = json!(184.4);
    assert_eq!(header_table_fragments(table, false), [(1, 0, 4)]);
}

#[test]
fn a_fitting_shared_start_keeps_the_legacy_header_decision() {
    let ends = table_rows(&[
        (1, json!({}), json!({"isHeader": true})),
        (1, json!({}), json!({})),
    ]);
    let keep = json!({"keepNext": true});
    let mut table = two_cell_table(&[(6, 0.0, keep.clone()), (6, 4.4, keep)]);
    let (lead, lead_measure) = paragraph(20, 2, json!({}));
    let (tail, tail_measure) =
        paragraph(21, 4, json!({"keepNext": true, "spacing": {"before": 4.4}}));
    table["block"]["rows"][0]["cells"][1]["blocks"] = json!([lead, tail]);
    table["measure"]["rows"][0]["cells"][1]["blocks"] = json!([lead_measure, tail_measure]);
    for key in ["block", "measure"] {
        table[key]["rows"] = json!([
            ends[key]["rows"][0],
            table[key]["rows"][0],
            ends[key]["rows"][1]
        ]);
    }
    table["measure"]["totalHeight"] = json!(164.4);
    assert_eq!(header_table_fragments(table, true), [(0, 0, 1), (1, 1, 3)]);
}

#[test]
fn keep_lines_in_one_cell_does_not_block_another_cells_split() {
    let measured = after_filler(
        2,
        two_cell_table(&[(6, 0.0, json!({})), (4, 4.4, json!({"keepLines": true}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 0.0)]);
    assert_eq!(
        placed_cell_lines(&measured),
        [vec![(0, 0..3)], vec![(0, 3..6), (1, 0..4)]]
    );
}

#[test]
fn per_cell_oversized_kept_remainders_yield_to_lines_and_terminate() {
    let measured = after_filler(
        2,
        two_cell_table(&[(6, 0.0, json!({})), (12, 4.4, json!({"keepLines": true}))]),
    );
    let result = table_fragments(measured.clone(), None);
    assert_eq!(result.len(), 3);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 60.0), (0.0, 44.4)]);
    assert!(result.iter().all(|(_, fragment)| {
        cell_windows(fragment)
            .iter()
            .any(|(top, bottom)| bottom > top)
    }));
    let lines = placed_cell_lines(&measured);
    for (cell, count) in [(0, 6), (1, 12)] {
        let seen: Vec<_> = lines
            .iter()
            .flatten()
            .filter(|(index, _)| *index == cell)
            .flat_map(|(_, lines)| lines.clone())
            .collect();
        assert_eq!(seen, (0..count).collect::<Vec<_>>());
    }
}

#[test]
fn an_unfinished_cell_can_progress_beyond_an_existing_shared_cut() {
    let mut table = two_cell_table(&[(8, 0.0, json!({})), (6, 0.0, json!({}))]);
    let (first, first_measure) = paragraph(11, 2, json!({}));
    let (last, last_measure) = paragraph(12, 4, json!({"spacing": {"before": 4.4}}));
    table["block"]["rows"][0]["cells"][1]["blocks"] = json!([first, last]);
    table["measure"]["rows"][0]["cells"][1]["blocks"] = json!([first_measure, last_measure]);
    table["measure"]["rows"][0]["cells"][1]["height"] = json!(124.4);
    let result = table_fragments(after_filler(1, table), None);
    assert_eq!(result[0].0, 0);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 80.0), (0.0, 40.0)]);
    assert_eq!(cell_windows(&result[1].1), [(80.0, 160.0), (40.0, 124.4)]);
}

#[test]
fn a_keep_next_heading_stays_with_a_per_cell_leading_slice() {
    let table = two_cell_table(&[(4, 0.0, json!({})), (3, 4.4, json!({}))]);
    assert_eq!(heading_and_table_pages_after(2, table), (Some(0), Some(0)));
}

#[test]
fn a_keep_next_row_stays_with_a_per_cell_leading_slice() {
    let heading = table_rows(&[(1, json!({"keepNext": true}), json!({}))]);
    let mut table = two_cell_table(&[(4, 0.0, json!({})), (3, 4.4, json!({}))]);
    table["block"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, heading["block"]["rows"][0].clone());
    table["measure"]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, heading["measure"]["rows"][0].clone());
    table["measure"]["totalHeight"] = json!(100);
    let result = table_fragments(after_filler(2, table), None);
    assert_eq!(result.len(), 2);
    assert_eq!(result[0].0, 0);
    assert_eq!(result[0].1["rowStart"], 0);
    assert_eq!(result[0].1["rowEnd"], 2);
    assert_eq!(result[0].1["height"], 60);
    assert_eq!(cell_windows(&result[0].1), [(0.0, 40.0), (0.0, 0.0)]);
}
