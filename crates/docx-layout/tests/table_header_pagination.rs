use serde_json::{Value, json};

fn paragraph(id: usize, heights: &[u32], keep_next: bool) -> Value {
    let lines: Vec<_> = heights
        .iter()
        .map(|height| {
            json!({
                "headRun":0,"headChar":0,"tailRun":0,"tailChar":1,
                "width":10,"ascent":8,"descent":2,"lineHeight":height
            })
        })
        .collect();
    json!({
        "block":{"kind":"paragraph","id":id,"runs":[{"kind":"text","text":"x"}],
            "attrs":{"keepNext":keep_next}},
        "measure":{"kind":"paragraph","lines":lines,"totalHeight":heights.iter().sum::<u32>()}
    })
}

fn table(headers: usize, row_lines: &[&[u32]]) -> Value {
    let mut rows = Vec::new();
    let mut measures = Vec::new();
    for (index, heights) in row_lines.iter().enumerate() {
        let paragraph = paragraph(index + 10, heights, false);
        let height = heights.iter().sum::<u32>();
        rows.push(json!({
            "id":index,"isHeader":index < headers,
            "cells":[{"id":index,"blocks":[paragraph["block"]]}]
        }));
        measures.push(json!({
            "height":height,"cells":[{"width":100,"height":height,"blocks":[paragraph["measure"]]}]
        }));
    }
    json!({
        "block":{"kind":"table","id":"table","rows":rows,"columnWidths":[100]},
        "measure":{"kind":"table","rows":measures,"columnWidths":[100],"totalWidth":100,
            "totalHeight":row_lines.iter().flat_map(|row| row.iter()).sum::<u32>()}
    })
}

fn layout(measured: Vec<Value>) -> Value {
    let input = json!({
        "measured":measured,
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10}}
    });
    serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
        .unwrap()
}

#[test]
fn leading_headers_move_with_the_first_body_slice() {
    let result = layout(vec![
        paragraph(1, &[50], false),
        table(2, &[&[20], &[20], &[40]]),
    ]);
    assert_eq!(result["pages"].as_array().unwrap().len(), 2);
    assert_eq!(result["pages"][0]["fragments"].as_array().unwrap().len(), 1);
    let fragment = &result["pages"][1]["fragments"][0];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 3);
    assert_eq!(fragment["height"], 80);
    assert!(fragment["headerRowCount"].is_null());
}

#[test]
fn first_body_slice_does_not_require_the_entire_splittable_row() {
    let result = layout(vec![
        paragraph(1, &[40], false),
        table(2, &[&[20], &[20], &[20, 20, 20, 20, 20]]),
    ]);
    let fragment = &result["pages"][0]["fragments"][1];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 3);
    assert_eq!(fragment["clipBottom"], 20);
    assert_eq!(fragment["height"], 60);
    assert_eq!(result["pages"][1]["fragments"][0]["headerRowCount"], 2);
}

#[test]
fn an_unsplittable_first_body_row_moves_with_the_headers() {
    let mut table = table(2, &[&[20], &[20], &[20, 20, 20]]);
    table["block"]["rows"][2]["cantSplit"] = json!(true);
    let result = layout(vec![paragraph(1, &[20], false), table]);
    assert_eq!(result["pages"].as_array().unwrap().len(), 2);
    assert_eq!(result["pages"][0]["fragments"].as_array().unwrap().len(), 1);
    assert_eq!(result["pages"][1]["fragments"][0]["height"], 100);
}

#[test]
fn oversized_header_bands_finish_without_repeating_the_partial_header() {
    for header_count in [2, 3] {
        let result = layout(vec![table(
            header_count,
            &[&[20, 20, 20], &[20, 20, 20], &[20]],
        )]);
        assert_eq!(result["pages"].as_array().unwrap().len(), 2);
        let first = &result["pages"][0]["fragments"][0];
        let second = &result["pages"][1]["fragments"][0];
        assert_eq!(first["rowEnd"], 2);
        assert_eq!(first["clipBottom"], 40);
        assert_eq!(second["rowStart"], 1);
        assert_eq!(second["clipTop"], 40);
        assert_eq!(second["rowEnd"], 3);
        assert_eq!(second["height"], 40);
        assert!(second["headerRowCount"].is_null());
    }
}

#[test]
fn header_metadata_is_omitted_when_the_band_cannot_repeat_with_the_body() {
    let result = layout(vec![table(2, &[&[40], &[40], &[40]])]);
    assert_eq!(result["pages"].as_array().unwrap().len(), 2);
    let fragment = &result["pages"][1]["fragments"][0];
    assert_eq!(fragment["rowStart"], 2);
    assert_eq!(fragment["height"], 40);
    assert!(fragment["headerRowCount"].is_null());
}

#[test]
fn all_header_tables_that_fit_move_as_one_band() {
    let result = layout(vec![paragraph(1, &[50], false), table(2, &[&[40], &[40]])]);
    assert_eq!(result["pages"].as_array().unwrap().len(), 2);
    assert_eq!(result["pages"][0]["fragments"].as_array().unwrap().len(), 1);
    assert_eq!(result["pages"][1]["fragments"][0]["height"], 80);
}

#[test]
fn a_keep_next_heading_moves_with_the_header_band_and_body_start() {
    let result = layout(vec![
        paragraph(1, &[40], false),
        paragraph(2, &[20], true),
        table(2, &[&[20], &[20], &[20]]),
    ]);
    assert_eq!(result["pages"].as_array().unwrap().len(), 2);
    assert_eq!(result["pages"][0]["fragments"].as_array().unwrap().len(), 1);
    assert_eq!(result["pages"][1]["fragments"][0]["blockId"], 2);
    assert_eq!(result["pages"][1]["fragments"][1]["rowEnd"], 3);
}
