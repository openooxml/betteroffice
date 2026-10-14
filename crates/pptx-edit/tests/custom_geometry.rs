use std::collections::BTreeMap;

use pptx_edit::DeckSession;
use pptx_parse::{PptxPackage, ShapeNode};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/custom-geometry.pptx");

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

#[test]
fn custom_paths_persist_without_changing_legacy_json_or_source_parts() {
    let session = DeckSession::open(FIXTURE, 285).unwrap();
    let json = serde_json::to_value(session.package()).unwrap();
    let shapes = json["slides"][0]["shapes"].as_array().unwrap();
    assert_eq!(shapes[0]["paths"].as_array().unwrap().len(), 4);
    assert!(shapes[0]["paths"][0].get("noFill").is_none());
    assert!(shapes[0]["paths"][0].get("noStroke").is_none());
    assert_eq!(shapes[0]["paths"][1]["noFill"], true);
    assert_eq!(shapes[0]["paths"][2]["noStroke"], true);
    assert!(shapes[4].get("paths").is_none());
    let mut legacy = json.clone();
    for slide in legacy["slides"].as_array_mut().unwrap() {
        for shape in slide["shapes"].as_array_mut().unwrap() {
            shape.as_object_mut().unwrap().remove("paths");
        }
    }
    let package: PptxPackage = serde_json::from_value(legacy.clone()).unwrap();
    assert_eq!(serde_json::to_value(&package).unwrap(), legacy);
    let restored =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 286).unwrap();
    assert_eq!(serde_json::to_value(restored.package()).unwrap(), json);
    assert_eq!(session.snapshot().unwrap(), restored.snapshot().unwrap());
    assert_eq!(parts(&session.save().unwrap()), parts(FIXTURE));
    let with_source = DeckSession::open_from_update_with_source(
        &restored.encode_state_as_update_v1(),
        FIXTURE,
        287,
    )
    .unwrap();
    assert_eq!(parts(&with_source.save().unwrap()), parts(FIXTURE));
    let ShapeNode::Shape(shape) = &with_source.package().slides[0].shapes[0] else {
        panic!()
    };
    assert_eq!(shape.paths.len(), 4);
}
