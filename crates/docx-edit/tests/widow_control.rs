//! Widow control keeps a paragraph's last two lines together: where only its
//! last line would carry over, Word moves one more line to the next page. The
//! fixtures fill an A4 page of 58 exact 12pt lines to leave three or four
//! lines for a four- or five-line paragraph (one wrapped by width, one after a
//! keepNext heading); their page breaks were exported from Word 16.113.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// The first and last text line on each page.
fn page_lines(name: &str) -> Vec<(String, String)> {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/widow-control")
            .join(format!("{name}.docx")),
    )
    .unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76600);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    });
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            let lines: Vec<String> = page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| {
                    fragment["resolvedLines"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                })
                .map(|line| {
                    line["segments"]
                        .as_array()
                        .into_iter()
                        .flatten()
                        .filter_map(|segment| segment["text"].as_str())
                        .collect::<String>()
                })
                .filter(|text| !text.is_empty())
                .collect();
            (lines[0].clone(), lines.last().unwrap().clone())
        })
        .collect()
}

#[test]
fn a_lone_last_line_takes_the_one_above_it_to_the_next_page() {
    for (name, last_on_page_one, first_on_page_two) in [
        ("four-lines-room-three", "L2", "L3"),
        ("heading-five-lines-room-four", "L3", "L4"),
    ] {
        let pages = page_lines(name);
        assert_eq!(pages.len(), 2, "{name}");
        assert_eq!(pages[0].1, last_on_page_one, "{name}");
        assert_eq!(pages[1].0, first_on_page_two, "{name}");
    }
    let pages = page_lines("wrapped-four-lines-room-three");
    assert!(pages[0].1.starts_with("word010"), "{pages:?}");
    assert!(pages[1].0.starts_with("word019"), "{pages:?}");
}
