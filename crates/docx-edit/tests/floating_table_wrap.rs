use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const WORD: &str = include_str!("fixtures/floating-table-wrap/word.json");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/floating-table-wrap")
            .join(format!("{name}.docx")),
    )
    .unwrap()
}

fn request(bytes: &[u8]) -> String {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections: Vec<_> = package
        .document
        .sections
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": sections, "settings": package.settings},
        "measurement": {
            "fontChains": {"arial|0|0": [font], "arial|1|0": [font]},
            "defaults": {"fontFamily": "Arial", "fontSize": 11}
        }
    })
    .to_string()
}

fn metrics(bytes: &[u8]) -> (f64, usize) {
    let engine = EngineSession::new(76600);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    let measured = output["measured"].as_array().unwrap();
    let anchor = measured
        .iter()
        .find(|block| block["block"]["kind"] == "paragraph")
        .expect("anchor paragraph");
    let anchor_lines = anchor["measure"]["lines"].as_array().unwrap().len();
    let marker = measured
        .iter()
        .find(|block| {
            block["block"]["kind"] == "paragraph"
                && block["block"]["runs"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|run| run["text"] == "MARKER HEADING")
        })
        .expect("MARKER HEADING paragraph");
    let lines = marker["measure"]["lines"].as_array().unwrap();
    assert_eq!(lines.len(), 1, "MARKER HEADING must fit on one line");
    for page in output["layout"]["pages"].as_array().unwrap() {
        for fragment in page["fragments"].as_array().unwrap() {
            if fragment["blockId"] == marker["block"]["id"] {
                assert_eq!(page["number"], 1, "MARKER HEADING must stay on page one");
                let y = fragment["y"].as_f64().unwrap()
                    + lines[0]["floatSkipBefore"].as_f64().unwrap_or(0.0);
                return (y * 72.0 / 96.0, anchor_lines);
            }
        }
    }
    panic!("no MARKER HEADING line");
}

fn assert_matches_word(name: &str) {
    let (marker_y, anchor_lines) = metrics(&fixture(name));
    let word: Value = serde_json::from_str(WORD).unwrap();
    let expected_y = word[name]["markerY"]
        .as_f64()
        .unwrap_or_else(|| {
            panic!("{name}: fill markerY in floating-table-wrap/word.json from Word")
        });
    let expected_lines = word[name]["anchorLines"]
        .as_u64()
        .unwrap_or_else(|| {
            panic!("{name}: fill anchorLines in floating-table-wrap/word.json from Word")
        });
    assert!(
        (marker_y - expected_y).abs() <= 2.0,
        "{name}: MARKER HEADING y={marker_y:.3}pt, Word={expected_y:.3}pt"
    );
    assert_eq!(anchor_lines as u64, expected_lines, "{name}: anchor lines");
}

#[test]
fn right_float_below_half_the_column_matches_word() {
    assert_matches_word("right-4660");
}

#[test]
fn right_float_just_above_half_the_column_matches_word() {
    assert_matches_word("right-4695");
}

#[test]
fn wide_right_float_matches_word() {
    assert_matches_word("right-5500");
}

#[test]
fn wide_left_float_matches_word() {
    assert_matches_word("left-5500");
}

#[test]
fn wide_right_float_with_an_empty_anchor_matches_word() {
    assert_matches_word("right-5500-empty-anchor");
}
