#![cfg(feature = "base-latin")]

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::structured::{ExportFailureCode, PageExportOptions, RevisionView};
use docx_edit::{EngineSession, seed_from_docx};
use ooxml_text::FontId;
use serde_json::{Value, json};

#[test]
fn backstop_measurements_remeasure_when_the_first_real_font_arrives() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = fixture::unrevised_docx();
    let engine = EngineSession::new(9665);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let backstop = FontId::base_latin(false, false).to_u32();
    let mut request = fixture::region_request(&engine, &bytes, backstop);
    let before: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let real = docx_layout::register_measure_font_bytes(fixture::OTHER_FONT).unwrap();
    assert_eq!(real, 0);
    for chain in request["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = json!([real]);
    }
    let warm: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    assert_ne!(before["measured"], warm["measured"]);
    let cold: Value = docx_layout::with_private_measure_fonts(|| {
        assert_eq!(
            docx_layout::register_measure_font_bytes(fixture::OTHER_FONT).unwrap(),
            0
        );
        let cold = EngineSession::new(9665);
        seed_from_docx(cold.doc(), &bytes).unwrap();
        serde_json::from_str(
            &cold
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap()
    });
    assert_eq!(warm, cold);
}

#[test]
fn backstop_only_layout_is_unavailable_for_paged_export() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = fixture::unrevised_docx();
    let engine = EngineSession::new(9666);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request =
        fixture::region_request(&engine, &bytes, FontId::base_latin(false, false).to_u32());
    let options = PageExportOptions::new(RevisionView::Markup);
    for ids in [
        vec![FontId::base_latin(false, false).to_u32()],
        vec![
            FontId::base_latin(true, false).to_u32(),
            FontId::base_latin(false, false).to_u32(),
        ],
    ] {
        for chain in request["measurement"]["fontChains"]
            .as_object_mut()
            .unwrap()
            .values_mut()
        {
            *chain = json!(ids);
        }
        engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap();
        let failure = engine.export_structured_with_pages(&options).unwrap_err();
        assert_eq!(failure.failure.code, ExportFailureCode::LayoutUnavailable);
        assert!(
            failure
                .failure
                .message
                .contains("No registered font measures")
        );
    }
    let real = docx_layout::register_measure_font_bytes(fixture::FONT).unwrap();
    for chain in request["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = json!([real]);
    }
    engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    engine.export_structured_with_pages(&options).unwrap();
}
