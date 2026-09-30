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

/// Per page, the table fragments' row ranges and heights in points.
fn table_fragments(name: &str) -> Vec<Vec<(u64, u64, f64)>> {
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
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|fragment| fragment["kind"] == "table")
                .map(|fragment| {
                    (
                        fragment["rowStart"].as_u64().unwrap(),
                        fragment["rowEnd"].as_u64().unwrap(),
                        fragment["height"].as_f64().unwrap() * 0.75,
                    )
                })
                .collect()
        })
        .collect()
}

fn pages(breaks: &[(&str, &str)]) -> Vec<(String, String)> {
    breaks
        .iter()
        .map(|(first, last)| ((*first).to_owned(), (*last).to_owned()))
        .collect()
}

/// A keepNext heading before a table needs room for the first row's smallest
/// slice only: one line, two under widow control, the whole row when it cannot
/// split. The heading has 1-3 lines of room above an 8-line (2-line) row.
#[test]
fn a_heading_keeps_with_the_first_slice_of_a_table_row() {
    for (name, heading_stays) in [
        ("keep-next-table-row-line-fits", true),
        ("keep-next-table-short-row-line-fits", true),
        ("keep-next-table-row-widow-fits", true),
        ("keep-next-table-row-widow-moves", false),
        ("keep-next-table-row-cant-split-moves", false),
    ] {
        let lines = page_lines(name);
        assert_eq!(lines.len(), 2, "{name}");
        assert_eq!(lines[0].1 == "H", heading_stays, "{name}: {lines:?}");
        assert_eq!(lines[1].0 == "H", !heading_stays, "{name}: {lines:?}");
    }
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

/// A row whose last paragraph keeps with the next row needs only the next
/// row's first line below it: with three lines of room, Word puts the
/// two-line first row and the first of the second row's four lines on page 1.
#[test]
fn a_row_keeps_with_the_first_line_of_the_next_row() {
    assert_eq!(
        table_fragments("keep-next-row-first-line-fits"),
        vec![vec![(0, 2, 36.0)], vec![(1, 2, 36.0)]]
    );
}

#[test]
fn a_page_break_before_the_follower_leaves_the_heading_in_place() {
    // 57 lines, a keepNext heading, then a paragraph with pageBreakBefore
    assert_eq!(
        page_lines("keep-next-page-break-before"),
        pages(&[("F01", "H"), ("B", "B")])
    );
}
