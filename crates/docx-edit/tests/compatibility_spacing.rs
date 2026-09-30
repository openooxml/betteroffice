use docx_edit::{EngineSession, bridge::RenderEnv, seed_from_docx};
use docx_layout::{
    cell_layout::layout_cell_content,
    types::{BlockExtent, LayoutBlock, MeasuredBlock},
};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const AUTO_DEFAULT: &[u8] = include_bytes!("fixtures/compatibility-spacing/auto-default.docx");
const AUTO_ENABLED: &[u8] = include_bytes!("fixtures/compatibility-spacing/auto-enabled.docx");
const AUTO_WITHOUT_FALLBACK: &[u8] =
    include_bytes!("fixtures/compatibility-spacing/auto-without-fallback.docx");
const BREAK_DEFAULT: &[u8] = include_bytes!("fixtures/compatibility-spacing/break-default.docx");
const BREAK_ENABLED: &[u8] = include_bytes!("fixtures/compatibility-spacing/break-enabled.docx");
const TABLE_DEFAULT: &[u8] = include_bytes!("fixtures/compatibility-spacing/table-default.docx");
const TABLE_ENABLED: &[u8] = include_bytes!("fixtures/compatibility-spacing/table-enabled.docx");

fn session(bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(76402);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
}

fn lowered(bytes: &[u8]) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    serde_json::from_str(
        &session(bytes)
            .lower_story_json(
                "body",
                &RenderEnv {
                    compatibility_flags: package.settings.compatibility_flags,
                    ..RenderEnv::default()
                },
            )
            .unwrap(),
    )
    .unwrap()
}

fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    serde_json::from_str(
        &session(bytes)
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                    "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn fragment(output: &Value, target: &str) -> (u64, f64) {
    for page in output["layout"]["pages"].as_array().unwrap() {
        for fragment in page["fragments"].as_array().unwrap() {
            let text: String = fragment["resolvedLines"]
                .as_array()
                .into_iter()
                .flatten()
                .flat_map(|line| line["segments"].as_array().into_iter().flatten())
                .filter_map(|segment| segment["text"].as_str())
                .collect();
            if text == target {
                return (
                    page["number"].as_u64().unwrap(),
                    fragment["y"].as_f64().unwrap(),
                );
            }
        }
    }
    panic!("missing {target} fragment");
}

fn word_position(fixture: &str, target: &str) -> (u64, f64) {
    let baseline: Value =
        serde_json::from_str(include_str!("fixtures/compatibility-spacing/word.json")).unwrap();
    for page in baseline[fixture].as_array().unwrap() {
        for line in page["lines"].as_array().unwrap() {
            if line["text"] == target {
                return (
                    page["index"].as_u64().unwrap() + 1,
                    line["y"].as_f64().unwrap(),
                );
            }
        }
    }
    panic!("missing {fixture}: {target} Word baseline");
}

fn close(actual: f64, expected: f64) {
    assert!((actual - expected).abs() < 0.2, "{actual} != {expected}");
}

#[test]
fn auto_spacing_preserves_the_default_and_uses_fixed_asymmetric_spacing_when_enabled() {
    for (bytes, before, after) in [
        (AUTO_DEFAULT, 14.0, 14.0),
        (AUTO_ENABLED, 100.0 / 15.0, 200.0 / 15.0),
        (AUTO_WITHOUT_FALLBACK, 100.0 / 15.0, 200.0 / 15.0),
    ] {
        let blocks = lowered(bytes);
        close(
            blocks[1]["attrs"]["spacing"]["before"].as_f64().unwrap(),
            before,
        );
        close(
            blocks[2]["attrs"]["spacing"]["after"].as_f64().unwrap(),
            after,
        );
        close(
            blocks[4]["attrs"]["spacing"]["before"].as_f64().unwrap(),
            before,
        );
        close(
            blocks[4]["attrs"]["spacing"]["after"].as_f64().unwrap(),
            after,
        );
    }
    let output = render(AUTO_ENABLED);
    for (previous, next) in [
        ("ANCHOR", "AUTO BEFORE"),
        ("AUTO AFTER", "AFTER TARGET"),
        ("AFTER TARGET", "AUTO BOTH"),
        ("AUTO BOTH", "BOTH TARGET"),
    ] {
        let (_, before_y) = fragment(&output, previous);
        let (page, after_y) = fragment(&output, next);
        assert_eq!(page, 1);
        close(
            after_y - before_y,
            (word_position("auto-enabled", next).1 - word_position("auto-enabled", previous).1)
                * 4.0
                / 3.0,
        );
    }
}

#[test]
fn hard_page_break_suppression_matches_word_and_preserves_page_break_before() {
    for (bytes, fixture) in [
        (BREAK_DEFAULT, "break-default"),
        (BREAK_ENABLED, "break-enabled"),
    ] {
        let output = render(bytes);
        let word_top = word_position(fixture, "FILLER").1;
        assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 4);
        for target in ["HARD TARGET", "PROPERTY TARGET", "STANDALONE TARGET"] {
            let (page, y) = fragment(&output, target);
            let (word_page, word_y) = word_position(fixture, target);
            assert_eq!(page, word_page);
            close(y, 96.0 + (word_y - word_top) * 4.0 / 3.0);
        }
    }
    let blocks = lowered(BREAK_ENABLED);
    assert_eq!(blocks[1]["attrs"]["pageBreakBeforeRun"], true);
    assert_eq!(blocks[1]["attrs"]["pageBreakBefore"], Value::Null);
    assert_eq!(blocks[1]["attrs"]["spacing"]["before"], 0.0);
}

#[test]
fn table_compatibility_preserves_unflagged_measurement_and_allows_cell_spacing() {
    let default = render(TABLE_DEFAULT);
    let enabled = render(TABLE_ENABLED);
    let default_blocks = lowered(TABLE_DEFAULT);
    let enabled_blocks = lowered(TABLE_ENABLED);
    for index in [0, 1] {
        assert_eq!(
            default_blocks[index]["attrs"]["spacing"],
            enabled_blocks[index]["attrs"]["spacing"]
        );
        assert_eq!(enabled_blocks[index]["attrs"]["contextualSpacing"], true);
    }
    for index in 0..3 {
        assert_eq!(
            default_blocks[2]["rows"][0]["cells"][0]["blocks"][index]["attrs"]["contextualSpacing"],
            true
        );
        assert_eq!(
            enabled_blocks[2]["rows"][0]["cells"][0]["blocks"][index]["attrs"]["contextualSpacing"],
            false
        );
    }
    for target in ["BODY 1", "BODY 2"] {
        assert_eq!(fragment(&default, target), fragment(&enabled, target));
    }
    let default_measure = &default["measured"][2]["measure"];
    assert_eq!(default_measure["totalHeight"], 96.0);
    assert_eq!(default_measure["rows"][0]["height"], 96.0);
    for (index, height) in [40.0, 56.0, 32.0].into_iter().enumerate() {
        assert_eq!(
            default_measure["rows"][0]["cells"][0]["blocks"][index]["totalHeight"],
            height
        );
    }
    assert_eq!(fragment(&default, "TABLE END"), (1, 240.0));
    assert_eq!(default_measure, &enabled["measured"][2]["measure"]);
    for (output, fixture) in [(&default, "table-default"), (&enabled, "table-enabled")] {
        let measured: MeasuredBlock =
            serde_json::from_value(output["measured"][2].clone()).unwrap();
        let (LayoutBlock::Table(table), BlockExtent::Table(extent)) =
            (&measured.block, &measured.measure)
        else {
            panic!("expected measured table");
        };
        let cell = layout_cell_content(
            Some(&table.rows[0].cells[0].blocks),
            Some(&extent.rows[0].cells[0].blocks),
            0.0,
        );
        for (index, (previous, next)) in [("CELL 1", "CELL 2"), ("CELL 2", "CELL 3")]
            .into_iter()
            .enumerate()
        {
            close(
                cell.line_tops[index + 1][0] - cell.line_tops[index][0],
                (word_position(fixture, next).1 - word_position(fixture, previous).1) * 4.0 / 3.0,
            );
        }
    }
    close(
        fragment(&enabled, "TABLE END").1 - fragment(&enabled, "BODY 1").1,
        (word_position("table-enabled", "TABLE END").1
            - word_position("table-enabled", "BODY 1").1)
            * 4.0
            / 3.0,
    );
}
