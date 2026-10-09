//! Word breaks a table row across pages only inside content that overflows
//! it: a row its `w:trHeight` minimum sizes moves whole to the next page
//! when it does not fit, even where its lines would. Rows whose content is
//! taller than the minimum break at a line as before. Page placement was
//! exported from Word 16.113 (`word.json`).

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture_dir() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/table-row-min-height")
}

fn request(bytes: &[u8]) -> String {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string()
}

#[test]
fn a_row_its_minimum_height_sizes_moves_whole_to_the_next_page() {
    let word: Value =
        serde_json::from_str(&std::fs::read_to_string(fixture_dir().join("word.json")).unwrap())
            .unwrap();
    for (name, expected) in word.as_object().unwrap() {
        let Some(expected) = expected["testLinesOnPageOne"].as_u64() else {
            continue;
        };
        let bytes = std::fs::read(fixture_dir().join(format!("{name}.docx"))).unwrap();
        let engine = EngineSession::new(76700);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let output: Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request(&bytes))
                .unwrap(),
        )
        .unwrap();
        let first_page_table = output["layout"]["pages"][0]["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .any(|fragment| fragment["kind"] == "table");
        assert_eq!(first_page_table, expected > 0, "{name}");
    }
}
