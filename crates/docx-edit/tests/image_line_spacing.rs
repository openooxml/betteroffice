//! Word draws an inline image that is alone on its line at the line's top and
//! adds below it the room an `auto` (multiple) rule adds to the paragraph
//! mark's single line: half a 12pt Arial line at 1.5, a whole one at 2. The
//! positions were measured in Word 16.113's PDF export (`word.json`).

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture_dir() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/image-line-spacing")
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
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string()
}

/// Each body fragment's top and each painted image's top on the first page, in points.
fn tops(bytes: &[u8]) -> (Vec<f64>, Vec<f64>) {
    let engine = EngineSession::new(76600);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_retained_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    let fragments = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .map(|fragment| fragment["y"].as_f64().unwrap() * 0.75)
        .collect();
    engine.build_display_list_frame("{}", 0).unwrap();
    let display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    let images = display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| primitive["kind"] == "image")
        .map(|image| image["y"].as_f64().unwrap() * 0.75)
        .collect();
    (fragments, images)
}

fn word() -> Value {
    serde_json::from_str(&std::fs::read_to_string(fixture_dir().join("word.json")).unwrap())
        .unwrap()
}

fn points(value: &Value) -> Vec<f64> {
    value
        .as_array()
        .unwrap()
        .iter()
        .map(|top| top.as_f64().unwrap())
        .collect()
}

fn assert_near(actual: &[f64], expected: &[f64], what: &str) {
    assert_within(actual, expected, 0.3, what);
}

fn assert_within(actual: &[f64], expected: &[f64], tolerance: f64, what: &str) {
    assert_eq!(actual.len(), expected.len(), "{what}: {actual:?}");
    for (actual, expected) in actual.iter().zip(expected) {
        assert!(
            (actual - expected).abs() < tolerance,
            "{what}: {actual} vs {expected}"
        );
    }
}

#[test]
fn an_image_alone_on_its_line_takes_the_multiple_rules_added_room_below_it() {
    let word = word();
    for name in ["img-body-240", "img-body-360", "img-body-480"] {
        let (fragments, images) =
            tops(&std::fs::read(fixture_dir().join(format!("{name}.docx"))).unwrap());
        let expected = points(&word[name]["imageTops"]);
        // Top A, image, Bottom A, image, Bottom B
        assert_eq!(fragments.len(), 5, "{name}: {fragments:?}");
        assert_near(&[fragments[1], fragments[3]], &expected, name);
        assert_near(&images, &expected, name);
        assert_near(
            &fragments[4..],
            &[word[name]["bottomBTop"].as_f64().unwrap()],
            name,
        );
    }
}

#[test]
fn an_image_alone_on_its_line_in_a_table_cell_takes_the_added_room_too() {
    let word = &word()["img-cell-300"];
    let (fragments, images) =
        tops(&std::fs::read(fixture_dir().join("img-cell-300.docx")).unwrap());
    assert_near(&images, &points(&word["imageTops"]), "image");
    // Top A, the table, Bottom A; the table's outer borders count half a
    // border less than in Word, independently of the image line.
    assert_within(
        &fragments[2..],
        &[word["bottomATop"].as_f64().unwrap()],
        0.7,
        "after the table",
    );
}
