use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn percentage_table_cell_margins_match_word_row_pitch() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for (mode, bytes, word_width, word_pitch, word_height) in [
        (
            14,
            include_bytes!("fixtures/table-outer-borders/percentage-cell-margins-14.docx")
                .as_slice(),
            123.60,
            39.5,
            71.0,
        ),
        (
            15,
            include_bytes!("fixtures/table-outer-borders/percentage-cell-margins-15.docx")
                .as_slice(),
            120.45,
            51.5,
            83.0,
        ),
    ] {
        let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
            .unwrap()
            .document
            .package;
        let request = json!({
            "bodyStory": "body", "renderEnv": {},
            "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
            "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 10}}
        });
        let engine = EngineSession::new(76801);
        seed_from_docx(engine.doc(), bytes).unwrap();
        engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap();
        let input: Value =
            serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
        let table = input["measured"]
            .as_array()
            .unwrap()
            .iter()
            .find(|block| block["block"]["kind"] == "table")
            .unwrap();
        let measure = &table["measure"];
        let pitch = measure["rows"][1]["height"].as_f64().unwrap() * 0.75;
        assert!(
            (pitch - word_pitch).abs() <= 0.15,
            "mode {mode}: middle row {pitch:.3}pt, Word {word_pitch:.3}pt"
        );
        let height = measure["totalHeight"].as_f64().unwrap() * 0.75;
        assert!(
            (height - word_height).abs() <= 0.15,
            "mode {mode}: table {height:.3}pt, Word {word_height:.3}pt"
        );
        let width = measure["columnWidths"][1].as_f64().unwrap() * 0.75;
        assert!(
            (width - word_width).abs() <= 0.05,
            "mode {mode}: column {width:.3}pt, Word {word_width:.3}pt"
        );
    }
}
