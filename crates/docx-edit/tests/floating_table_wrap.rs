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
    let layout = engine
        .layout_document_with_regions_json(&request(bytes))
        .unwrap();
    let output: Value = serde_json::from_str(&layout).unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&layout).unwrap()).unwrap();
    let anchor_lines = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|block| block["block"]["kind"] == "paragraph")
        .expect("anchor paragraph")["measure"]["lines"]
        .as_array()
        .unwrap()
        .len();
    let runs: Vec<&Value> = display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .collect();
    let mut text = String::new();
    let mut starts = Vec::new();
    for run in &runs {
        starts.push(text.len());
        text.push_str(run["text"].as_str().unwrap_or_default());
    }
    let at = text
        .find("MARKER HEADING")
        .unwrap_or_else(|| panic!("MARKER HEADING on page one, page text {text:?}"));
    let marker = runs[starts.partition_point(|&start| start <= at) - 1];
    let baseline = marker["baselineY"]
        .as_f64()
        .or_else(|| marker["glyphs"][0]["y"].as_f64())
        .unwrap();
    (baseline * 0.75, anchor_lines)
}

fn assert_matches_word(name: &str) {
    let (marker_baseline, anchor_lines) = metrics(&fixture(name));
    let word: Value = serde_json::from_str(WORD).unwrap();
    let expected_baseline = word[name]["markerBaseline"].as_f64().unwrap();
    let expected_lines = word[name]["anchorLines"].as_u64().unwrap();
    assert!(
        (marker_baseline - expected_baseline).abs() <= 2.0,
        "{name}: MARKER HEADING baseline {marker_baseline:.2}pt, Word {expected_baseline}pt"
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

#[test]
fn wide_right_float_anchored_to_the_margin_matches_word() {
    assert_matches_word("right-4820-margin");
}

#[test]
fn wide_right_float_without_a_horizontal_anchor_matches_word() {
    assert_matches_word("right-4820-no-anchor");
}
