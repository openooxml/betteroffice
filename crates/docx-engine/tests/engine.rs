#![cfg(not(target_arch = "wasm32"))]

use std::sync::Arc;

use docx_engine::{Document, EngineErrorKind, Host};
use serde_json::Value;

fn synthetic_docx() -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        ),
    ])
    .unwrap()
}

fn document() -> Document {
    Document::new(
        7,
        Host::new(Arc::new(|| 0), Arc::new(|| 1), Arc::new(|| 0.0)),
    )
}

fn assert_paragraphs(document: &Document) {
    let paragraphs: Value = serde_json::from_str(
        &document
            .read_paragraphs_json(r#"{"story":"body","view":"accepted"}"#)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(paragraphs["ok"], true);
    assert_eq!(paragraphs["paragraphs"][0]["text"], "Engine text");
    assert_eq!(paragraphs["paragraphs"].as_array().unwrap().len(), 1);
    assert!(!document.encode_state().is_empty());
}

#[test]
fn document_opens_and_seeds_without_wasm() {
    let bytes = synthetic_docx();
    let opened = document();
    opened.open_docx(&bytes, true, None, None).unwrap();
    assert_paragraphs(&opened);

    let seeded = document();
    seeded.seed_from_docx(&bytes, None).unwrap();
    assert_paragraphs(&seeded);

    let loaded = document();
    loaded.load(&opened.encode_state()).unwrap();
    assert_paragraphs(&loaded);
}

#[test]
fn parse_s9_returns_json() {
    let bytes = synthetic_docx();
    let parsed: Value =
        serde_json::from_str(&docx_engine::parse_docx_s9(&bytes, "").unwrap()).unwrap();
    assert!(parsed["wireVersion"].is_u64());
    assert_eq!(
        parsed["document"]["package"]["document"]["content"]
            .as_array()
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn zip_entries_round_trip() {
    let bytes = synthetic_docx();
    let entries = docx_engine::unzip_docx(&bytes).unwrap();
    assert_eq!(entries.len(), 3);
    let rezipped = docx_engine::rezip_docx(entries.clone()).unwrap();
    assert_eq!(docx_engine::unzip_docx(&rezipped).unwrap(), entries);
}

#[test]
fn display_list_queries_share_the_home_registry() {
    let json = r#"{"pages":[{"pageIndex":0,"width":100,"height":100,"primitives":[]}]}"#;
    let handle = docx_engine::open_display_list(json).unwrap();
    let expected = docx_engine::hit_test_regions_json(json, 0, 20.0, 20.0).unwrap();
    assert_eq!(
        docx_layout::session::hit_test_regions_by_handle(handle, 0, 20.0, 20.0).unwrap(),
        expected
    );
    assert_eq!(
        docx_engine::hit_test_regions_by_handle(handle, 0, 20.0, 20.0).unwrap(),
        expected
    );
    docx_engine::close_display_list(handle);
    let error = docx_engine::hit_test_regions_by_handle(handle, 0, 20.0, 20.0).unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(
        error.message,
        format!("unknown display-list handle {handle}")
    );
}

#[test]
fn invalid_page_windows_return_before_parsing_or_handle_lookup() {
    assert_eq!(
        docx_engine::range_rects_on_pages_json("invalid", 0.0, 1.0, f64::NAN, 0.0).unwrap(),
        "[]"
    );
    assert_eq!(
        docx_engine::range_rects_on_pages_by_handle(0, 0.0, 1.0, f64::NAN, 0.0).unwrap(),
        "[]"
    );
}

#[test]
fn free_fonts_work_beside_an_unentered_document() {
    let _document = document();
    docx_engine::clear_measure_fonts().unwrap();
    assert!(docx_engine::register_measure_font(&[]).is_err());
    assert!(docx_engine::outline_glyph_json(0, 0).is_err());
    docx_engine::clear_measure_fonts().unwrap();
}

#[test]
fn parse_errors_keep_the_home_message() {
    let expected = docx_parse::parse_docx_s9_json(&[], "{").unwrap_err();
    let error = docx_engine::parse_docx_s9(&[], "{").unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(error.message, expected);
}
