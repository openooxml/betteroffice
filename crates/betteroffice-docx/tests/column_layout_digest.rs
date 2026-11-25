mod common;

use betteroffice_docx::{Document, LayoutInput};
use serde_json::{Value, json};

fn input(columns: Value) -> LayoutInput {
    let mut value: Value = serde_json::from_str(include_str!(
        "../../docx-layout/tests/fixtures/single-page-multi-paragraph.input.json"
    ))
    .unwrap();
    value["options"]["columns"] = columns;
    serde_json::from_value(value).unwrap()
}

#[test]
fn column_layout_digest() {
    let document = Document::open(&common::sample_docx()).unwrap();
    for columns in [
        json!({"count": 2, "gap": 36}),
        json!({"count": 3, "gap": 24, "separator": true}),
        json!({"count": 2, "gap": 36, "equalWidth": false,
               "columns": [{"width": 200}, {"width": 160}]}),
        json!({"count": 3, "gap": 24, "equalWidth": false, "separator": true}),
    ] {
        let result = document.layout(input(columns.clone())).unwrap();
        let layout = serde_json::to_string(&result.layout).unwrap();
        let display = serde_json::to_string(&result.display_list).unwrap();
        println!(
            "DIGEST {columns} {} {}",
            layout.len(),
            fnv(&layout) ^ fnv(&display).rotate_left(1)
        );
    }
}

fn fnv(text: &str) -> u64 {
    text.bytes().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(byte)).wrapping_mul(0x0100_0000_01b3)
    })
}
