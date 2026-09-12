use serde_json::{Value, json};

fn input() -> Value {
    let mut input: Value = serde_json::from_str(include_str!(
        "fixtures/table-splits-with-repeated-header.input.json"
    ))
    .unwrap();
    for row in input["measured"][0]["block"]["rows"]
        .as_array_mut()
        .unwrap()
    {
        row["cantSplit"] = json!(true);
    }
    input
}

fn layout(input: &Value) -> Value {
    serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
        .unwrap()
}

fn keep_next(input: &mut Value, row: usize) {
    input["measured"][0]["block"]["rows"][row]["cells"][0]["blocks"][0]["attrs"] =
        json!({"keepNext": true});
}

#[test]
fn a_heading_row_moves_with_the_following_row() {
    let mut input = input();
    keep_next(&mut input, 2);
    let result = layout(&input);
    assert_eq!(result["pages"][0]["fragments"][0]["rowEnd"], 2);
    assert_eq!(result["pages"][0]["fragments"][0]["height"], 168);
    assert_eq!(result["pages"][1]["fragments"][0]["rowStart"], 2);
    assert_eq!(result["pages"][1]["fragments"][0]["rowEnd"], 4);
}

#[test]
fn a_heading_and_follower_that_fit_stay_on_the_current_page() {
    let mut input = input();
    let expected = layout(&input);
    keep_next(&mut input, 1);
    assert_eq!(layout(&input), expected);
}

#[test]
fn oversized_keep_chains_do_not_create_blank_pages() {
    let mut input = input();
    let expected = layout(&input);
    keep_next(&mut input, 1);
    keep_next(&mut input, 2);
    assert_eq!(layout(&input), expected);
}

#[test]
fn a_kept_table_group_can_move_to_the_next_column() {
    let mut input = input();
    input["options"]["columns"] = json!({"count": 2, "gap": 24});
    keep_next(&mut input, 2);
    let result = layout(&input);
    let fragments = result["pages"][0]["fragments"].as_array().unwrap();
    assert_eq!(fragments.len(), 2);
    assert_eq!(fragments[0]["rowEnd"], 2);
    assert_eq!(fragments[1]["rowStart"], 2);
    assert!(fragments[1]["x"].as_f64().unwrap() > fragments[0]["x"].as_f64().unwrap());
}

#[test]
fn a_table_cell_page_break_does_not_force_a_row_break() {
    let mut input = input();
    let expected = layout(&input);
    input["measured"][0]["block"]["rows"][2]["cells"][0]["blocks"][0]["attrs"] =
        json!({"pageBreakBefore": true});
    assert_eq!(layout(&input), expected);
}
