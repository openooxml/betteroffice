use docx_edit::{EngineSession, seed_from_docx};
use docx_layout::resolve_lines::resolve_line_segments;
use docx_layout::types::{BlockExtent, LayoutBlock, TableBlock, TableExtent};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn check(name: &str, bytes: &[u8]) {
    docx_layout::with_private_measure_fonts(|| {
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(78901);
        seed_from_docx(engine.doc(), bytes).unwrap();
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
                "fontChains": {"arial|0|0": [font]},
                "defaults": {"fontFamily": "Arial", "fontSize": 12},
                "authoritativeShaping": true
            }
        });
        let output: Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let oracle: Value =
            serde_json::from_str(include_str!("fixtures/table-widths/word.json")).unwrap();
        let expected = &oracle[name];
        let pages = output["layout"]["pages"].as_array().unwrap();
        let starts: Vec<Value> = pages
            .iter()
            .map(|page| page["fragments"][0]["rowStart"].clone())
            .collect();
        assert_eq!(
            json!(starts),
            expected["pageStartRows"],
            "{name}: page boundaries"
        );
        let block: TableBlock =
            serde_json::from_value(output["measured"][0]["block"].clone()).unwrap();
        let extent: TableExtent =
            serde_json::from_value(output["measured"][0]["measure"].clone()).unwrap();
        for (row_index, (row, measured)) in block.rows.iter().zip(&extent.rows).enumerate() {
            for (cell_index, (cell, measured)) in row.cells.iter().zip(&measured.cells).enumerate()
            {
                let LayoutBlock::Paragraph(paragraph) = &cell.blocks[0] else {
                    panic!("paragraph")
                };
                let BlockExtent::Paragraph(extent) = &measured.blocks[0] else {
                    panic!("paragraph extent")
                };
                let lines: Vec<String> = extent
                    .lines
                    .iter()
                    .map(|line| {
                        assert_ne!(line.synthetic_fallback, Some(true));
                        resolve_line_segments(&paragraph.runs, line)
                            .iter()
                            .map(|segment| segment.text.as_str())
                            .collect::<String>()
                            .trim()
                            .to_owned()
                    })
                    .collect();
                let expected_lines: Vec<String> = expected["cellLines"][cell_index]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|line| {
                        line.as_str()
                            .unwrap()
                            .replace("001", &format!("{:03}", row_index + 1))
                    })
                    .collect();
                assert_eq!(
                    lines, expected_lines,
                    "{name}: row {row_index} cell {cell_index}"
                );
            }
        }
    });
}

macro_rules! fixture {
    ($test:ident, $name:literal) => {
        #[test]
        fn $test() {
            check(
                $name,
                include_bytes!(concat!("fixtures/table-widths/", $name, ".docx")),
            );
        }
    };
}

fixture!(fixed_preferred_widths_match_word, "fixed-preferred");
fixture!(fixed_spanning_cell_matches_word, "fixed-span");
fixture!(autofit_matches_word, "autofit");
fixture!(autofit_spanning_cell_matches_word, "autofit-span");
fixture!(omitted_layout_defaults_to_autofit, "default-autofit");
fixture!(nowrap_matches_word, "nowrap");
fixture!(margins_and_indentation_match_word, "margins-indent");
