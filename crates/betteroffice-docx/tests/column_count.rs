mod common;

use std::sync::mpsc::channel;
use std::thread;
use std::time::Duration;

use betteroffice_docx::{Document, LayoutInput};
use docx_layout::types::{ColumnLayout, MAX_COLUMNS};
use serde_json::{Value, json};

fn columns(count: f64) -> ColumnLayout {
    ColumnLayout {
        count,
        gap: 36.0,
        equal_width: Some(false),
        separator: None,
        columns: None,
    }
}

fn input(columns: Option<ColumnLayout>) -> LayoutInput {
    let mut value: Value = serde_json::from_str(include_str!(
        "../../docx-layout/tests/fixtures/single-page-multi-paragraph.input.json"
    ))
    .unwrap();
    for block in value["measured"].as_array_mut().unwrap() {
        if let Some(runs) = block["block"]["runs"].as_array_mut() {
            for run in runs {
                run["fontFamily"] = json!("Calibri");
            }
        }
    }
    let mut input: LayoutInput = serde_json::from_value(value).unwrap();
    input.options.columns = columns;
    input
}

fn pages(columns: Option<ColumnLayout>) -> usize {
    let (sender, receiver) = channel();
    thread::spawn(move || {
        let document = Document::open(&common::sample_docx()).unwrap();
        let pages = document
            .layout(input(columns))
            .map(|result| result.layout.pages.len())
            .map_err(|error| error.to_string());
        let _ = sender.send(pages);
    });
    receiver
        .recv_timeout(Duration::from_secs(10))
        .expect("layout returns")
        .unwrap()
}

#[test]
fn column_count_is_clamped_to_the_laid_out_range() {
    for (count, expected) in [
        (2.0, 2),
        (0.0, 1),
        (-3.0, 1),
        (f64::NAN, 1),
        (45.0, MAX_COLUMNS),
        (1e9, MAX_COLUMNS),
        (f64::INFINITY, MAX_COLUMNS),
    ] {
        assert_eq!(columns(count).column_count(), expected, "{count}");
    }
}

#[test]
fn layout_returns_for_a_huge_unequal_column_count() {
    assert_eq!(pages(Some(columns(1e9))), pages(Some(columns(45.0))));
}

#[test]
fn layout_keeps_two_unequal_columns() {
    assert_eq!(pages(Some(columns(2.0))), 1);
}

#[test]
fn layout_returns_for_a_zero_column_count() {
    assert!(pages(Some(columns(0.0))) >= 1);
}

#[test]
fn layout_returns_for_a_negative_column_count() {
    assert!(pages(Some(columns(-3.0))) >= 1);
}

#[test]
fn layout_returns_for_a_nan_column_count() {
    assert!(pages(Some(columns(f64::NAN))) >= 1);
}

#[test]
fn layout_returns_for_an_infinite_column_count() {
    assert_eq!(pages(Some(columns(f64::INFINITY))), pages(Some(columns(45.0))));
}
