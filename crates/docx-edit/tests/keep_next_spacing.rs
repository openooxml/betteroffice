//! A keepNext paragraph moves to the next page only when its run and the
//! start of its follower cannot finish above the page bottom, where every gap
//! is the larger of the space-after above it and the space-before below it
//! and each paragraph's spacing counts once. The fixtures fill an A4 page of
//! 58 exact 12pt lines to within a few twips of the run's height; their page
//! breaks were exported from Word 16.113.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// The first and last text line on each page.
fn page_lines(name: &str) -> Vec<(String, String)> {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/keep-next")
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

fn pages(breaks: &[(&str, &str)]) -> Vec<(String, String)> {
    breaks
        .iter()
        .map(|(first, last)| ((*first).to_owned(), (*last).to_owned()))
        .collect()
}

#[test]
fn a_heading_counts_its_space_after_once() {
    assert_eq!(
        page_lines("keep-next-spacing-fits"),
        pages(&[("Line 001", "Follower two"), ("Line 055", "Line 080")])
    );
}

#[test]
fn a_heading_counts_only_the_larger_gap_above_it() {
    assert_eq!(
        page_lines("keep-next-collapsed-gap-fits"),
        pages(&[("Line 001", "Follower two"), ("Line 053", "Line 080")])
    );
}

#[test]
fn a_heading_still_moves_when_its_run_does_not_fit() {
    assert_eq!(
        page_lines("keep-next-collapsed-gap-moves"),
        pages(&[("Line 001", "Line 053"), ("Heading", "Line 080")])
    );
}

#[test]
fn a_chain_of_headings_collapses_the_gaps_between_them() {
    assert_eq!(
        page_lines("keep-next-chain-fits"),
        pages(&[("Line 001", "Line 052"), ("Line 053", "Line 080")])
    );
}
