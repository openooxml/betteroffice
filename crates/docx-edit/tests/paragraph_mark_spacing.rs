use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const SPACES: &[u8] = include_bytes!("fixtures/paragraph-mark-spacing/spaces.docx");
const FORMATTED_SPACES: &[u8] =
    include_bytes!("fixtures/paragraph-mark-spacing/formatted-spaces.docx");

fn setup(bytes: &[u8]) -> (EngineSession, Value) {
    let base = docx_layout::register_measure_font(FONT).unwrap();
    let font = docx_layout::register_substitute_measure_font(base, "ＭＳ ゴシック").unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {
            "sections": [{"properties": package.document.final_section_properties}],
            "settings": package.settings
        },
        "measurement": {
            "fontChains": {"ｍｓ ゴシック|0|0": [font]},
            "defaults": {"fontFamily": "ＭＳ ゴシック", "fontSize": 9}
        }
    });
    let engine = EngineSession::new(76900);
    seed_from_docx(engine.doc(), bytes).unwrap();
    (engine, request)
}

fn layout(bytes: &[u8]) -> (Value, Value) {
    let (engine, request) = setup(bytes);
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    (output, request)
}

fn tops(bytes: &[u8]) -> Vec<f64> {
    let (output, _) = layout(bytes);
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .map(|fragment| fragment["y"].as_f64().unwrap() * 0.75)
        .collect()
}

fn assert_tops(bytes: &[u8], expected: &[f64]) {
    let actual = tops(bytes);
    assert_eq!(actual.len(), expected.len());
    assert!(
        (actual.last().unwrap() - expected.last().unwrap()).abs() < 0.4,
        "following paragraph: {:.3}pt, Word {:.3}pt",
        actual.last().unwrap(),
        expected.last().unwrap()
    );
    for (index, (actual, expected)) in actual.iter().zip(expected).enumerate() {
        assert!(
            (actual - expected).abs() < 0.6,
            "paragraph {index}: {actual:.3}pt, Word {expected:.3}pt"
        );
    }
}

#[test]
fn oversized_spaces_use_the_paragraph_mark_for_line_height() {
    assert_tops(SPACES, &[72.0, 84.55, 97.10, 109.65, 122.25, 134.80]);
    assert_tops(
        FORMATTED_SPACES,
        &[72.0, 84.55, 97.10, 109.65, 122.25, 134.80],
    );
}

#[test]
fn visible_text_keeps_its_run_size_for_line_height() {
    assert_tops(
        include_bytes!("fixtures/paragraph-mark-spacing/visible.docx"),
        &[72.0, 84.55, 98.80, 112.20, 126.40, 140.65],
    );
}

fn unpatched_measure(output: &Value, request: &Value, index: usize) -> Value {
    let block = &output["measured"][index]["block"];
    let fragment = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .find(|fragment| fragment["blockId"] == block["id"])
        .unwrap();
    let mut input = request["measurement"].clone();
    input["block"] = block.clone();
    input["maxWidth"] = fragment["width"].clone();
    input["authoritativeShaping"] = json!(true);
    serde_json::from_str(&docx_layout::measure_paragraph_json_resident(&input.to_string()).unwrap())
        .unwrap()
}

fn horizontal_geometry(line: &Value) -> Value {
    let mut line = line.clone();
    for key in ["lineHeight", "ascent", "descent"] {
        line.as_object_mut().unwrap().remove(key);
    }
    line
}

#[test]
fn spacer_lines_preserve_unpatched_run_ranges_and_widths() {
    let (output, request) = layout(FORMATTED_SPACES);
    for index in [3, 4] {
        let original = unpatched_measure(&output, &request, index);
        assert_eq!(original["lines"][0]["tailRun"], 1);
        assert_eq!(original["lines"][0]["tailChar"], 1);
        assert!(original["lines"][0]["width"].as_f64().unwrap() > 0.0);
        let lines = output["measured"][index]["measure"]["lines"]
            .as_array()
            .unwrap();
        assert_eq!(lines.len(), 1);
        assert_eq!(
            horizontal_geometry(&lines[0]),
            horizontal_geometry(&original["lines"][0]),
            "spacer {index}: run geometry differs from unpatched measurement"
        );
    }
}

#[test]
fn carets_and_hits_resolve_inside_spacer_runs() {
    let (output, _) = layout(FORMATTED_SPACES);
    let display: Value = serde_json::from_str(
        &docx_layout::display_list::build_display_list_json(&output.to_string()).unwrap(),
    )
    .unwrap();
    let typed: docx_layout::display_list::DisplayList =
        serde_json::from_value(display.clone()).unwrap();
    for index in [3, 4] {
        let start = output["measured"][index]["block"]["runs"][0]["pmStart"]
            .as_f64()
            .unwrap() as i64;
        let position = start + 1;
        let caret = docx_layout::hit::caret_rect(&typed, position)
            .unwrap_or_else(|| panic!("spacer {index}: no caret inside spaces at {position}"));
        let first = docx_layout::hit::caret_rect(&typed, start).unwrap();
        assert!(caret.x > first.x);
        assert_eq!(
            docx_layout::hit::hit_test(
                &typed,
                caret.page_index,
                caret.x,
                caret.y + caret.height / 2.0,
            ),
            Some(position)
        );
    }
    let primitives = display["pages"][0]["primitives"].as_array().unwrap();
    assert!(
        primitives
            .iter()
            .any(|primitive| primitive["kind"] == "decoration"
                && primitive["deco"] == "highlight"
                && primitive["w"].as_f64().unwrap() > 0.0)
    );
    assert!(
        primitives
            .iter()
            .any(|primitive| primitive["kind"] == "decoration"
                && primitive["deco"] == "underline"
                && primitive["w"].as_f64().unwrap() > 0.0)
    );
}

#[test]
fn grid_spacers_with_run_snap_disabled_never_grow() {
    let (output, request) = layout(include_bytes!("fixtures/paragraph-mark-spacing/grid.docx"));
    for index in [3, 4] {
        let original = unpatched_measure(&output, &request, index);
        let block = &output["measured"][index]["block"];
        assert!(block["attrs"]["docGridPitchPx"].as_f64().unwrap() > 30.0);
        assert_eq!(block["runs"][0]["snapToGrid"], false);
        let measure = &output["measured"][index]["measure"];
        let height = measure["totalHeight"].as_f64().unwrap();
        let original_height = original["totalHeight"].as_f64().unwrap();
        assert!(
            height <= original_height,
            "grid spacer {index}: {height:.3}px exceeds main {original_height:.3}px"
        );
        assert_eq!(measure["lines"], original["lines"]);
    }
}

#[test]
fn retained_spacer_metrics_match_fresh_layout_after_an_edit() {
    let (engine, request) = setup(FORMATTED_SPACES);
    let request = request.to_string();
    let initial = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let inputs: Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    let (fresh, _) = layout(FORMATTED_SPACES);
    for index in 1..5 {
        assert_eq!(
            inputs["measured"][index]["measure"],
            fresh["measured"][index]["measure"]
        );
    }
    assert_eq!(
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap(),
        initial
    );
    let position = inputs["measured"][3]["block"]["runs"][0]["pmStart"]
        .as_f64()
        .unwrap() as i64
        + 1;
    engine
        .doc()
        .insert_text(
            &docx_edit::EditCtx::local("", ""),
            docx_edit::Position::new("body", position.try_into().unwrap()),
            " ",
            docx_edit::FormatPolicy::Inherit,
        )
        .unwrap();
    let before = engine.stats();
    let incremental = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before.incremental_pagination_calls + 1
    );
    let fresh = EngineSession::new(76901);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh
            .layout_document_with_regions_retained_json(&request)
            .unwrap(),
        incremental
    );
    let actual: Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    let expected: Value =
        serde_json::from_str(&fresh.retained_kernel_inputs_json().unwrap()).unwrap();
    assert_eq!(actual["measured"], expected["measured"]);
}
