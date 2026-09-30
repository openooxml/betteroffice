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
    let (block, measure) = paragraph(10, lines, attrs);
    let height = 20 * lines;
    json!({
        "block":{"kind":"table","id":"table","columnWidths":[100],
            "rows":[{"id":0,"cells":[{"id":0,"blocks":[block]}]}]},
        "measure":{"kind":"table","columnWidths":[100],"totalWidth":100,"totalHeight":height,
            "rows":[{"height":height,"cells":[{"width":100,"height":height,"blocks":[measure]}]}]}
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
