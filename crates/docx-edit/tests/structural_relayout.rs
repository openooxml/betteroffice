//! A structural relayout (Enter, paste) re-measures only the touched blocks:
//! the full region pass reuses retained measures by stable block key, which
//! survives paragraph splits.

use docx_edit::{EditCtx, EngineSession, Position, seed_from_docx};

const SENTENCE: &str = "The issuer shall maintain and make available to the competent authority a complete and accurate record of every crypto-asset service provided during the reporting period, including the identity of the client, the nature of the instruction and the timestamp at which it was executed.";

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

fn build_docx(paragraph_count: usize) -> Vec<u8> {
    let mut body = String::new();
    for i in 1..=paragraph_count {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">Section {i}. {SENTENCE}</w:t></w:r></w:p>"#
        ));
    }
    body.push_str(
        r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#,
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap()
}

#[test]
fn structural_relayout_reuses_retained_measures() {
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), &build_docx(400)).unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "properties": {
                "pageWidth": 12240, "pageHeight": 15840,
                "marginTop": 1440, "marginRight": 1440,
                "marginBottom": 1440, "marginLeft": 1440
            }
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font_id]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    })
    .to_string();
    engine.layout_document_with_regions_json(&request).unwrap();

    engine
        .doc()
        .split_paragraph(&EditCtx::local("", ""), Position::new("body", 700), None)
        .unwrap();
    let before = engine.stats();
    let retained = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let stats = engine.stats();

    // only the two halves of the split re-measure; every other block reuses
    // its retained measure by stable key
    assert_eq!(
        stats.resident_measure_calls - before.resident_measure_calls,
        2
    );
    assert_eq!(
        stats.resident_reused_blocks - before.resident_reused_blocks,
        399
    );
    assert_eq!(stats.retained_measured_blocks, 401);

    // the retained reply carries pagination only — no measured arena
    let value: serde_json::Value = serde_json::from_str(&retained).unwrap();
    assert!(value.get("measured").is_none());
    assert!(value.get("layout").is_some());
    assert_eq!(value["notesConverged"], true);

    // the measured arena stays fetchable on demand for the display fallback
    let kernel: serde_json::Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    assert_eq!(kernel["measured"].as_array().unwrap().len(), 401);
    assert!(kernel.get("options").is_some());
}
