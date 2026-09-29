//! A paragraph mark inside a complex field's code is part of the code, so
//! Word hides it and joins the paragraph with the next one. Each fixture
//! holds an IF field whose code spans three paragraph marks, followed by its
//! result, on A4 at an exact 12pt pitch. The page breaks were exported from
//! Word 16.113.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/field-code-paragraphs")
            .join(format!("{name}.docx")),
    )
    .unwrap()
}

/// Each page's body lines, as text.
fn pages(bytes: &[u8]) -> Vec<Vec<String>> {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
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
            page["fragments"]
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
                        .cloned()
                        .unwrap_or_default()
                        .iter()
                        .filter_map(|segment| segment["text"].as_str().map(str::to_owned))
                        .collect::<String>()
                })
                .collect()
        })
        .collect()
}

fn page_ranges(pages: &[Vec<String>]) -> Vec<(String, String)> {
    pages
        .iter()
        .map(|lines| (lines[0].clone(), lines.last().unwrap().clone()))
        .collect()
}

#[test]
fn a_header_field_code_over_paragraph_marks_takes_one_line() {
    let pages = pages(&fixture("header-field-code-paragraphs"));
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        [
            "Line 001", "Line 055", "Line 109", "Line 163", "Line 217", "Line 271"
        ]
    );
}

#[test]
fn a_body_field_code_over_paragraph_marks_joins_its_paragraphs() {
    let pages = pages(&fixture("body-field-code-paragraphs"));
    assert!(
        pages[0]
            .iter()
            .any(|line| line.trim_end() == "Line 041 yes")
    );
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"]
    );
}
