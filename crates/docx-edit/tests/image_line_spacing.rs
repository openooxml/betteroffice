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

/// Each body paragraph fragment's top on the first page, in points.
fn fragment_tops(bytes: &[u8]) -> Vec<f64> {
    let engine = EngineSession::new(76600);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .map(|fragment| fragment["y"].as_f64().unwrap() * 0.75)
        .collect()
}

#[test]
fn an_image_alone_on_its_line_takes_the_multiple_rules_added_room_below_it() {
    let word: Value =
        serde_json::from_str(&std::fs::read_to_string(fixture_dir().join("word.json")).unwrap())
            .unwrap();
    for name in ["img-body-240", "img-body-360", "img-body-480"] {
        let tops =
            fragment_tops(&std::fs::read(fixture_dir().join(format!("{name}.docx"))).unwrap());
        let expected = &word[name];
        let image_tops: Vec<f64> = expected["imageTops"]
            .as_array()
            .unwrap()
            .iter()
            .map(|top| top.as_f64().unwrap())
            .collect();
        // Top A, image, Bottom A, image, Bottom B
        assert_eq!(tops.len(), 5, "{name}: {tops:?}");
        for (actual, expected) in [tops[1], tops[3]].into_iter().zip(image_tops) {
            assert!((actual - expected).abs() < 0.3, "{name}: {tops:?}");
        }
        let bottom = expected["bottomBTop"].as_f64().unwrap();
        assert!((tops[4] - bottom).abs() < 0.3, "{name}: {tops:?}");
    }
}
