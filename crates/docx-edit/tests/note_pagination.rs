#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position};
use serde_json::Value;

#[test]
fn a_body_edit_beside_notes_repaginates_incrementally_as_a_fresh_layout_would() {
    let (engine, request) = fixture::laid_out(&fixture::unrevised_docx(), 9301);
    let reserved = |layout: &str| {
        serde_json::from_str::<Value>(layout).unwrap()["options"]["footnoteReservedHeights"].clone()
    };
    let original = reserved(&engine.layout_document_with_regions_json(&request).unwrap());
    assert!(
        original
            .as_object()
            .is_some_and(|heights| !heights.is_empty())
    );

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 2),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let before = engine.stats();
    let edited = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    assert_eq!(reserved(&edited), original);
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );

    let fresh = EngineSession::new(9302);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        edited
    );
}

fn floating_table(id: u32) -> String {
    format!(
        concat!(
            r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" "#,
            r#"w:vertAnchor="text" w:tblpY="1"/><w:tblW w:w="2000" w:type="dxa"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr>"#,
            r#"<w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#
        ),
        fixture::p(&format!("{id:08X}"), &fixture::r("Floating cell text"))
    )
}

#[test]
fn a_body_edit_beside_floats_remeasures_only_its_flow_segment_as_a_fresh_layout_would() {
    let words = "wrapped words beside a floating table ".repeat(8);
    let mut body = String::new();
    let mut id = 0x1000_u32;
    for segment in 0..3 {
        id += 1;
        let first = if segment == 0 {
            fixture::r(&words)
        } else {
            format!("<w:pPr><w:pageBreakBefore/></w:pPr>{}", fixture::r(&words))
        };
        body += &fixture::p(&format!("{id:08X}"), &first);
        id += 1;
        body += &floating_table(id);
        for _ in 0..4 {
            id += 1;
            body += &fixture::p(&format!("{id:08X}"), &fixture::r(&words));
        }
    }
    let bytes = fixture::with_body(&body);
    let (engine, request) = fixture::laid_out(&bytes, 9303);
    engine.layout_document_with_regions_json(&request).unwrap();

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 3),
            "longer ",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let before = engine.stats();
    let edited = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    let blocks = serde_json::from_str::<Value>(&edited).unwrap()["measured"]
        .as_array()
        .unwrap()
        .len() as u64;
    let remeasured = after.resident_measure_calls - before.resident_measure_calls;
    assert!(
        remeasured > 1 && remeasured < blocks / 2,
        "{remeasured} of {blocks}"
    );

    let fresh = EngineSession::new(9304);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        edited
    );
}
