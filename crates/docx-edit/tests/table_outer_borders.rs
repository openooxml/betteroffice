//! Word stacks one horizontal border above every row and adds the table's
//! bottom border after its last row, exact-height or not, so a bordered
//! table is a full border taller than the sum of its row pitches and the
//! first row's text sits a full top border below the table top.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn table_outer_borders_match_word_positions() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/table-outer-borders");
    let word: Value =
        serde_json::from_str(include_str!("fixtures/table-outer-borders/word.json")).unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for (name, expected) in word.as_object().unwrap() {
        let bytes = std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap();
        let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
            .unwrap()
            .document
            .package;
        let request = json!({
            "bodyStory": "body", "renderEnv": {},
            "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
            "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
        });
        let engine = EngineSession::new(76800);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let layout = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let output: Value = serde_json::from_str(&layout).unwrap();
        let pages = output["layout"]["pages"].as_array().unwrap();
        let paragraph_top = |page: usize, target: &str| {
            pages[page]["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .find(|fragment| {
                    let text: String = fragment["resolvedLines"]
                        .as_array()
                        .into_iter()
                        .flatten()
                        .flat_map(|line| line["segments"].as_array().into_iter().flatten())
                        .filter_map(|segment| segment["text"].as_str())
                        .collect();
                    fragment["kind"] == "paragraph" && text == target
                })
                .unwrap_or_else(|| panic!("{name}: missing {target} on page {}", page + 1))["y"]
                .as_f64()
                .unwrap()
                * 0.75
        };
        let top_a = expected["topA"].as_f64().unwrap();
        let bottom_page = expected["bottomA"][0].as_u64().unwrap() as usize;
        let word_bottom = expected["bottomA"][1].as_f64().unwrap() - top_a;
        assert_eq!(pages.len(), bottom_page, "{name}");
        let bottom = paragraph_top(bottom_page - 1, "Bottom A") - paragraph_top(0, "Top A");
        assert!(
            (bottom - word_bottom).abs() <= 0.15,
            "{name}: following paragraph {bottom:.3}pt, Word {word_bottom:.1}pt"
        );

        let display: Value =
            serde_json::from_str(&engine.build_display_list_json(&layout).unwrap()).unwrap();
        let mut lines: Vec<(f64, String)> = Vec::new();
        for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
            let (Some("text"), Some(text), Some(y)) = (
                primitive["kind"].as_str(),
                primitive["text"].as_str(),
                primitive["baselineY"].as_f64(),
            ) else {
                continue;
            };
            match lines.iter_mut().find(|(line_y, _)| *line_y == y) {
                Some((_, line)) => line.push_str(text),
                None => lines.push((y, text.to_owned())),
            }
        }
        let baseline = |target: &str| {
            lines
                .iter()
                .find(|(_, line)| line.contains(target))
                .unwrap_or_else(|| panic!("{name}: missing {target} text"))
                .0
                * 0.75
        };
        let row_text = expected["row01"][0].as_str().unwrap();
        let word_row = expected["row01"][1].as_f64().unwrap() - top_a;
        let row = baseline(row_text) - baseline("Top A");
        assert!(
            (row - word_row).abs() <= 0.15,
            "{name}: first row {row:.3}pt, Word {word_row:.1}pt"
        );
    }
}
