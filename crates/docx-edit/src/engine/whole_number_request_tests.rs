use super::{EngineSession, lowering_fixture::{NS, Package}};
use crate::seed_from_docx;
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn notes_document() -> Vec<u8> {
    let note = |tag: &str, id: u32, text: &str| {
        format!(r#"<w:{tag} w:id="{id}"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:{tag}>"#)
    };
    let body = concat!(
        r#"<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>"#,
        r#"<w:r><w:endnoteReference w:id="2"/></w:r></w:p>"#,
        r#"<w:sectPr><w:footnotePr><w:numStart w:val="8"/></w:footnotePr>"#,
        r#"<w:endnotePr><w:numStart w:val="2"/></w:endnotePr>"#,
        r#"<w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="1"/></w:sectPr>"#,
    );
    Package::new(body)
        .part(
            "footnotes.xml",
            "rIdFootnotes",
            "footnotes",
            "footnotes",
            &format!("<w:footnotes {NS}>{}</w:footnotes>", note("footnote", 1, "Footnote")),
        )
        .part(
            "endnotes.xml",
            "rIdEndnotes",
            "endnotes",
            "endnotes",
            &format!("<w:endnotes {NS}>{}</w:endnotes>", note("endnote", 2, "Endnote")),
        )
        .part(
            "settings.xml",
            "rIdSettings",
            "settings",
            "settings",
            &format!(
                r#"<w:settings {NS}><w:footnotePr><w:numStart w:val="1"/></w:footnotePr><w:endnotePr><w:numStart w:val="1"/></w:endnotePr></w:settings>"#
            ),
        )
        .bytes()
}

/// The request a host builds by forwarding the parsed model as JSON, with no
/// conversion of the values the model holds as `f64`.
fn forwarded_request(bytes: &[u8]) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut sections: Vec<Value> = package
        .document
        .sections
        .clone()
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"sectionId": section.id, "properties": section.properties}))
        .collect();
    sections.push(json!({"properties": package.document.final_section_properties}));
    let contents: Vec<Value> = [(&package.footnotes, "footnote"), (&package.endnotes, "endnote")]
        .into_iter()
        .flat_map(|(notes, kind)| {
            notes
                .iter()
                .flatten()
                .filter(|note| note.note_type.is_empty() || note.note_type == "normal")
                .map(move |note| json!({"id": note.id, "noteKind": kind, "height": 0}))
        })
        .collect();
    json!({
        "bodyStory": "body",
        "renderEnv": {},
        "options": {"pageGap": 24},
        "regions": {"sections": sections, "settings": package.settings},
        "notes": {"contents": contents},
    })
}

#[test]
fn a_forwarded_parse_model_lays_out_with_its_whole_number_floats() {
    let bytes = notes_document();
    let mut request = forwarded_request(&bytes);
    let text = request.to_string();
    for float in [r#""id":1.0"#, r#""numStart":8.0"#, r#""numStart":2.0"#, r#""numStart":1.0"#] {
        assert!(text.contains(float), "{float} missing from {text}");
    }
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(41);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let requirements: Vec<Value> =
        serde_json::from_str(&engine.layout_font_requirements_json(&text).unwrap()).unwrap();
    let chains: serde_json::Map<String, Value> = requirements
        .iter()
        .map(|requirement| (requirement["key"].as_str().unwrap().to_owned(), json!([font])))
        .collect();
    request["measurement"] = json!({
        "fontChains": chains,
        "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        "authoritativeShaping": true,
    });
    let layout: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    assert!(layout.is_object(), "{layout}");
}

#[test]
fn a_fractional_note_number_is_rejected_naming_the_field() {
    let engine = EngineSession::new(42);
    let mut request = forwarded_request(&notes_document());
    request["regions"]["settings"]["footnotePr"]["numStart"] = json!(1.5);
    let error = engine
        .layout_font_requirements_json(&request.to_string())
        .unwrap_err();
    assert!(error.contains("numStart must be an integer, got 1.5"), "{error}");
}
