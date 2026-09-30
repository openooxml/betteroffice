//! Word splits a table row only where every cell's paragraphs may break:
//! widow/orphan control (on unless `w:widowControl w:val="0"`) keeps a
//! paragraph's first and last two lines together, so a paragraph of two or
//! three lines never splits, and `w:keepLines` keeps it whole. Each fixture
//! fills an A4 page of 58 exact 12pt lines to leave one to three lines for a
//! two-cell row whose cells hold one multi-line paragraph each. The page
//! breaks were exported from Word 16.113.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn layout(name: &str) -> Value {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/table-row-widows")
            .join(format!("{name}.docx")),
    )
    .unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76700);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    });
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap()
}

/// Lines of the row's table on each page it reaches, as (page index, lines).
fn row_lines(name: &str) -> Vec<(usize, u32)> {
    layout(name)["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .flat_map(|(index, page)| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|fragment| fragment["kind"] == "table")
                .map(move |fragment| (index, (fragment["height"].as_f64().unwrap() / 16.0) as u32))
                .collect::<Vec<_>>()
        })
        .collect()
}

#[test]
fn a_two_line_paragraph_moves_its_row_whole() {
    assert_eq!(row_lines("row-two-line-widow"), [(1, 2)]);
}

#[test]
fn a_paragraph_without_widow_control_splits_after_one_line() {
    assert_eq!(row_lines("row-two-line-no-widow"), [(0, 1), (1, 1)]);
}

#[test]
fn a_three_line_paragraph_moves_its_row_whole() {
    assert_eq!(row_lines("row-three-line-widow"), [(1, 3)]);
}

#[test]
fn a_four_line_paragraph_splits_two_and_two() {
    assert_eq!(row_lines("row-four-line-widow"), [(0, 2), (1, 2)]);
    assert_eq!(row_lines("row-four-line-widow-three-fit"), [(0, 2), (1, 2)]);
}

#[test]
fn keep_lines_moves_the_row_whole() {
    assert_eq!(row_lines("row-four-line-keep-lines"), [(1, 4)]);
}

#[test]
fn every_cell_must_allow_the_break() {
    assert_eq!(row_lines("row-mixed-widow-one-fits"), [(1, 4)]);
}
