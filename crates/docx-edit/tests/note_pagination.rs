#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::collections::BTreeMap;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, ParaAttrDelta, ParaSelector, Position};
use serde_json::{Value, json};

fn extras(request: &str) -> String {
    let request: Value = serde_json::from_str(request).unwrap();
    json!({"fontChains": request["measurement"]["fontChains"]}).to_string()
}

fn display(engine: &EngineSession) -> Value {
    engine
        .with_display_list(|list| serde_json::to_value(list).unwrap())
        .unwrap()
}

/// Lays out and displays `engine`'s document in a fresh session with `request`.
fn fresh(engine: &EngineSession, request: &str, client_id: u64) -> (String, Value) {
    let fresh = EngineSession::new(client_id);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let layout = fresh.layout_document_with_regions_json(request).unwrap();
    fresh.build_display_list_frame(&extras(request), 0).unwrap();
    (layout, display(&fresh))
}

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

#[test]
fn a_body_edit_moves_note_backlinks_on_later_pages_as_a_fresh_layout_would() {
    let (engine, request) = fixture::laid_out(&fixture::unrevised_docx(), 9305);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 2),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    engine
        .build_display_list_frame(&extras(&request), 1)
        .unwrap();
    assert!(engine.stats().incremental_display_builds > 0);
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9306));
}

#[test]
fn removing_a_page_break_before_the_second_page_pulls_it_back_as_a_fresh_layout_would() {
    let body = format!(
        "{}{}{}",
        fixture::p(
            "20000001",
            &format!(
                r#"{}<w:r><w:footnoteReference w:id="1"/></w:r>"#,
                fixture::r("A short first page")
            )
        ),
        fixture::p(
            "20000002",
            &format!(
                "<w:pPr><w:pageBreakBefore/></w:pPr>{}",
                fixture::r("Second page")
            )
        ),
        fixture::p("20000003", &fixture::r("After it")),
    );
    let bytes = fixture::with_body_and_note(&body, &fixture::p("20000009", &fixture::r("Note")));
    let (engine, request) = fixture::laid_out(&bytes, 9307);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    let pages = |layout: &str| {
        serde_json::from_str::<Value>(layout).unwrap()["layout"]["pages"]
            .as_array()
            .unwrap()
            .len()
    };
    assert_eq!(
        pages(&engine.layout_document_with_regions_json(&request).unwrap()),
        2
    );
    engine
        .build_display_list_frame(&extras(&request), 1)
        .unwrap();

    engine
        .doc()
        .set_paragraph_attrs(
            &EditCtx::local("", ""),
            &ParaSelector::One("20000002".to_owned()),
            &ParaAttrDelta {
                other: BTreeMap::from([("pageBreakBefore".to_owned(), None)]),
                ..Default::default()
            },
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    engine
        .build_display_list_frame(&extras(&request), 2)
        .unwrap();
    assert_eq!(pages(&layout), 1);
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9308));
}

#[test]
fn narrowing_the_page_remeasures_floats_as_a_fresh_layout_would() {
    let words = "wrapped words beside a floating table ".repeat(8);
    let body = format!(
        "{}{}{}",
        fixture::p("30000001", &fixture::r(&words)),
        floating_table(0x3000_0002),
        fixture::p("30000003", &fixture::r(&words)),
    );
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9309);
    engine.layout_document_with_regions_json(&request).unwrap();
    let mut narrowed: Value = serde_json::from_str(&request).unwrap();
    for section in narrowed["regions"]["sections"].as_array_mut().unwrap() {
        section["properties"]["pageWidth"] = json!(7200);
        section["properties"]["marginLeft"] = json!(2400);
    }
    let narrowed = narrowed.to_string();
    let layout = engine.layout_document_with_regions_json(&narrowed).unwrap();
    engine
        .build_display_list_frame(&extras(&narrowed), 0)
        .unwrap();
    assert_ne!(
        layout,
        fresh(&engine, &request, 9310).0,
        "the narrowed request lays out differently"
    );
    assert_eq!((layout, display(&engine)), fresh(&engine, &narrowed, 9311));
}

#[test]
fn a_page_numbering_change_with_a_body_edit_restamps_pages_as_a_fresh_layout_would() {
    let (engine, request) = fixture::laid_out(&fixture::unrevised_docx(), 9312);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    let mut renumbered: Value = serde_json::from_str(&request).unwrap();
    for section in renumbered["regions"]["sections"].as_array_mut().unwrap() {
        if let Some(properties) = section["properties"].as_object_mut() {
            properties.remove("pageNumbering");
        }
    }
    let renumbered = renumbered.to_string();
    assert_ne!(renumbered, request, "the fixture numbers its pages");
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 2),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine
        .layout_document_with_regions_json(&renumbered)
        .unwrap();
    engine
        .build_display_list_frame(&extras(&renumbered), 1)
        .unwrap();
    assert_eq!(
        (layout, display(&engine)),
        fresh(&engine, &renumbered, 9313)
    );
}
