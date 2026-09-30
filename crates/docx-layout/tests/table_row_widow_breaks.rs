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

/// Lays out on 100px-tall pages and returns each table fragment's page and height.
fn fragments(measured: Vec<Value>) -> Vec<(usize, f64)> {
    let input = json!({"measured":measured,"options":{"pageSize":{"w":200,"h":120},
        "margins":{"top":10,"right":10,"bottom":10,"left":10}}});
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
                .map(move |fragment| (page, fragment["height"].as_f64().unwrap()))
                .collect::<Vec<_>>()
        })
        .collect()
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
fn a_kept_paragraph_taller_than_a_page_starts_on_a_fresh_page_and_splits() {
    let (block, measure) = paragraph(1, 2, json!({}));
    let above = json!({"block":block,"measure":measure});
    assert_eq!(
        fragments(vec![above, table(8, json!({"keepLines": true}))]),
        [(1, 100.0), (2, 60.0)]
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
        [(1, 100.0), (2, 100.0), (3, 60.0)]
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
