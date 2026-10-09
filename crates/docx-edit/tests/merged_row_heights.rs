use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {
            "sections": [{"properties": package.document.final_section_properties}],
            "settings": package.settings
        },
        "measurement": {
            "fontChains": {"arial|0|0": [font]},
            "defaults": {"fontFamily": "Arial", "fontSize": 12}
        }
    });
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap()
}

fn assert_word_fixture(name: &str, bytes: &[u8]) {
    let word: Value =
        serde_json::from_str(include_str!("fixtures/merged-row-heights/word.json")).unwrap();
    let expected = &word[name];
    let output = render(bytes);
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| measured["block"]["kind"] == "table")
        .unwrap();
    let source_rows = table["block"]["rows"].as_array().unwrap();
    for cell in &source_rows[0]["cells"].as_array().unwrap()[..2] {
        assert_eq!(cell["rowSpan"].as_f64(), Some(3.0), "{name}");
    }
    let measured_rows = table["measure"]["rows"].as_array().unwrap();
    for (index, height) in expected["mergedRowHeightsPt"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
    {
        let actual_pt = measured_rows[index]["height"].as_f64().unwrap() * 0.75;
        let expected_pt = height.as_f64().unwrap();
        assert!(
            (actual_pt - expected_pt).abs() < 0.3,
            "{name}, row {index}: {actual_pt}pt, Word {expected_pt}pt"
        );
    }
    for (index, row) in measured_rows.iter().enumerate().skip(3) {
        assert!(
            (row["height"].as_f64().unwrap() - 16.0).abs() < 0.01,
            "{name}, body row {index}"
        );
    }

    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(
        pages.len(),
        expected["pages"].as_array().unwrap().len(),
        "{name}"
    );
    let mut actual_windows = Vec::new();
    for page in pages {
        for fragment in page["fragments"].as_array().unwrap() {
            if fragment["kind"] != "table" {
                continue;
            }
            assert_eq!(fragment["blockId"], table["block"]["id"], "{name}");
            assert_eq!(fragment["clipTop"].as_f64().unwrap_or(0.0), 0.0, "{name}");
            assert_eq!(
                fragment["clipBottom"].as_f64().unwrap_or(0.0),
                0.0,
                "{name}"
            );
            actual_windows.push(json!({
                "page": page["number"],
                "rowStart": fragment["rowStart"],
                "rowEnd": fragment["rowEnd"],
                "repeatedHeaderRows": fragment["headerRowCount"].as_f64().unwrap_or(0.0) as u64
            }));
        }
    }
    assert_eq!(json!(actual_windows), expected["pages"], "{name}");
}

#[test]
fn same_span_cells_grow_to_the_tallest_requirement_in_either_column_order() {
    const CASES: [(&str, &[u8]); 2] = [
        (
            "same-span",
            include_bytes!("fixtures/merged-row-heights/same-span.docx"),
        ),
        (
            "same-span-reversed",
            include_bytes!("fixtures/merged-row-heights/same-span-reversed.docx"),
        ),
    ];
    for (name, bytes) in CASES {
        assert_word_fixture(name, bytes);
    }
}

#[test]
fn repeated_merged_headers_match_words_page_and_row_boundaries() {
    const CASES: [(&str, &[u8]); 2] = [
        (
            "repeated-header",
            include_bytes!("fixtures/merged-row-heights/repeated-header.docx"),
        ),
        (
            "repeated-header-reversed",
            include_bytes!("fixtures/merged-row-heights/repeated-header-reversed.docx"),
        ),
    ];
    for (name, bytes) in CASES {
        assert_word_fixture(name, bytes);
    }
}
