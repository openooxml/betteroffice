use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

#[path = "support/quality_fixture.rs"]
mod quality_fixture;
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn anchored_header_preserves_body_space_and_paints_at_page_coordinates() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(73002);
    seed_from_docx(engine.doc(), &quality_fixture::document(false)).unwrap();
    let request = json!({
        "bodyStory": "body",
        "options": {},
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 1122.6666666666667, "h": 793.3333333333334},
            "margins": {"top": 113.33333333333333, "right": 113.33333333333333,
                "bottom": 113.33333333333333, "left": 113.33333333333333,
                "header": 56.666666666666664, "footer": 56.666666666666664},
            "headerFooterRefs": {"headerFirst": "rIdHeader"}
        }]},
        "measurement": {
            "fontChains": {"arial|0|0": [font], "calibri|0|0": [font]},
            "defaults": {"fontSize": 12, "fontFamily": "Arial"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let value: Value = serde_json::from_str(&output).unwrap();
    assert_eq!(value["layout"]["pages"].as_array().unwrap().len(), 1);
    assert!(
        (value["options"]["margins"]["top"].as_f64().unwrap() - 113.33333333333333).abs() < 0.001
    );
    let variant = &value["headersFooters"]["variants"][0];
    assert!(variant["flowHeight"].as_f64().unwrap() < 30.0);
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let primitives = display["pages"][0]["header"]["primitives"]
        .as_array()
        .unwrap();
    let shape = primitives
        .iter()
        .find(|primitive| primitive["kind"] == "shape")
        .expect("header shape must paint");
    assert_eq!(shape["x"], 1010);
    assert!((shape["y"].as_f64().unwrap() - 56.667).abs() < 0.001);
    assert_eq!(shape["w"], 50);
}
