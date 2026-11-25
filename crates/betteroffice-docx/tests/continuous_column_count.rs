mod common;

use std::sync::mpsc::channel;
use std::thread;
use std::time::Duration;

use betteroffice_docx::{Document, LayoutInput};
use serde_json::{Value, json};

fn continuous_sections(count: f64) -> LayoutInput {
    let mut value: Value = serde_json::from_str(include_str!(
        "../../docx-layout/tests/fixtures/continuous-section-column-change.input.json"
    ))
    .unwrap();
    let columns = json!({"count": count, "gap": 24, "equalWidth": false});
    value["options"]["columns"] = columns.clone();
    value["options"]["bodyBreakType"] = json!("continuous");
    let margins = value["options"]["margins"].clone();
    for measured in value["measured"].as_array_mut().unwrap() {
        if measured["block"]["kind"] == "sectionBreak" {
            measured["block"]["type"] = json!("continuous");
            measured["block"]["columns"] = columns.clone();
            measured["block"]["margins"] = margins.clone();
        }
    }
    serde_json::from_value(value).unwrap()
}

fn continuous_pages(count: f64) -> usize {
    let (sender, receiver) = channel();
    thread::spawn(move || {
        let document = Document::open(&common::sample_docx()).unwrap();
        let pages = document
            .layout(continuous_sections(count))
            .map(|result| result.layout.pages.len())
            .map_err(|error| error.to_string());
        let _ = sender.send(pages);
    });
    receiver
        .recv_timeout(Duration::from_secs(30))
        .expect("layout returns")
        .unwrap()
}

#[test]
fn continuous_sections_with_matching_huge_column_counts_lay_out() {
    assert_eq!(continuous_pages(1e9), continuous_pages(45.0));
}
