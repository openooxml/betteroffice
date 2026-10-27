use super::{EngineSession, lowering_pages};
use crate::{StoryRange, bridge::RenderEnv, seed_from_docx};
use serde_json::Value;

fn document(text: &str) -> Vec<u8> {
    let body = format!(
        "{}{}<w:sectPr/>",
        lowering_pages::p("71000000", &lowering_pages::r("A")),
        lowering_pages::p(
            "71000001",
            &format!(
                r#"{}<w:r><w:br w:type="page"/></w:r>"#,
                if text.is_empty() {
                    String::new()
                } else {
                    lowering_pages::r(text)
                }
            ),
        ),
    );
    let mut parts = ooxml_opc::unzip_parts(&lowering_pages::with_body(&body)).unwrap();
    for (name, bytes) in &mut parts {
        let root = match name.as_str() {
            "word/footnotes.xml" => "footnotes",
            "word/endnotes.xml" => "endnotes",
            _ => continue,
        };
        *bytes = format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_edit_matches_fresh(before: &str, after: &str, text: Option<&str>, pages: usize) {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let bytes = document(before);
    let engine = EngineSession::new(9410);
    engine.set_local_lowering(true);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(
        engine.stats().retained_pages,
        if before.is_empty() { 1 } else { 2 }
    );
    engine
        .edit_resident_text(
            StoryRange::new("body", 2, 2 + before.len() as u32),
            text,
            true,
        )
        .unwrap();
    engine
        .layout_document_with_regions_after_edit(&request, Some("body"))
        .unwrap();
    let actual: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
    assert_eq!(actual["layout"]["pages"].as_array().unwrap().len(), pages);

    let fresh = EngineSession::new(9411);
    seed_from_docx(fresh.doc(), &document(after)).unwrap();
    let expected: Value =
        serde_json::from_str(&fresh.layout_document_with_regions_json(&request).unwrap()).unwrap();
    assert_eq!(actual["layout"], expected["layout"]);
    assert_eq!(
        engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap(),
        fresh
            .lower_story_json("body", &RenderEnv::default())
            .unwrap()
    );
}

#[test]
fn typing_into_a_terminal_break_only_paragraph_matches_fresh_open() {
    assert_edit_matches_fresh("", "x", Some("x"), 2);
}

#[test]
fn emptying_a_terminal_text_and_break_paragraph_matches_fresh_open() {
    assert_edit_matches_fresh("x", "", None, 1);
}
