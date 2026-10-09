#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::collections::BTreeMap;

use docx_edit::frame_delta::{FRAME_HEADER_LEN, PAGE_OP_LEN};
use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, ParaAttrDelta, ParaSelector, Position, StoryRange,
    UndoSession,
};
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

fn bulk_laid_out(bytes: &[u8], client_id: u64) -> (EngineSession, String) {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let seed = docx_edit::EditingDoc::new(client_id + 1);
    docx_edit::seed_from_docx(&seed, bytes).unwrap();
    let engine = EngineSession::new(client_id);
    engine.layout_document_with_regions_retained("{}").unwrap();
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    engine
        .doc()
        .set_note_separator_state(seed.note_separator_state().unwrap());
    let request = fixture::region_request(&engine, bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    (engine, request)
}

fn bulk_edit(engine: &EngineSession, step: Value) {
    let request = serde_json::from_value(json!({
        "expectVersion": engine.doc().version(),
        "history": "none",
        "steps": [step],
    }))
    .unwrap();
    assert!(
        engine
            .doc()
            .apply_edits(&request, &UndoSession::new())
            .unwrap()
            .unwrap()
            .applied
    );
}

/// Lays out and displays `engine`'s document in a fresh session with `request`.
fn fresh(engine: &EngineSession, request: &str, client_id: u64) -> (String, Value) {
    let fresh = EngineSession::new(client_id);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    fresh
        .doc()
        .set_note_separator_state(engine.doc().note_separator_state().unwrap());
    let layout = fresh.layout_document_with_regions_json(request).unwrap();
    fresh.build_display_list_frame(&extras(request), 0).unwrap();
    (layout, display(&fresh))
}

#[test]
fn synthetic_note_sections_typing_and_preview_toggles_match_cold_bytes() {
    let body: String = (0..16)
        .map(|index| {
            let mut content = if index == 7 {
                r#"<w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr>"#.to_owned()
            } else {
                String::new()
            };
            content += &fixture::r(&format!("Editable paragraph {index}"));
            if index == 2 {
                content += r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#;
            }
            if index == 5 {
                content += r#"<w:ins w:id="10" w:author="BetterOffice" w:date="2026-10-01T00:00:00Z"><w:r><w:t> pending</w:t></w:r></w:ins>"#;
            }
            fixture::p(&format!("{:08X}", 0x7100_0000 + index), &content)
        })
        .collect();
    let bytes = fixture::with_body_and_note(
        &small_page(&body),
        &fixture::p("71000100", &fixture::r("A synthetic footnote")),
    );
    let (engine, request) = fixture::laid_out(&bytes, 9380);
    let mut request: Value = serde_json::from_str(&request).unwrap();
    assert!(request["regions"]["sections"].as_array().unwrap().len() >= 2);
    engine
        .build_display_list_frame(&extras(&request.to_string()), 0)
        .unwrap();
    let assert_cold = |request: &Value| {
        let (cold, cold_display) = fresh(&engine, &request.to_string(), 9381);
        let cold: Value = serde_json::from_str(&cold).unwrap();
        let retained: Value =
            serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
        let inputs: Value =
            serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
        for (actual, expected) in [
            (&retained["layout"], &cold["layout"]),
            (&inputs["measured"], &cold["measured"]),
            (&inputs["options"], &cold["options"]),
        ] {
            assert_eq!(
                serde_json::to_vec(actual).unwrap(),
                serde_json::to_vec(expected).unwrap()
            );
        }
        assert_eq!(
            serde_json::to_vec(&display(&engine)).unwrap(),
            serde_json::to_vec(&cold_display).unwrap()
        );
    };
    assert_cold(&request);
    for decision in ["accepted", "rejected"] {
        for text in [Some("x"), Some("y"), None] {
            match text {
                Some(text) => engine
                    .doc()
                    .insert_text(
                        &EditCtx::local("", ""),
                        Position::new("body", 3),
                        text,
                        FormatPolicy::Inherit,
                    )
                    .map(|_| ()),
                None => engine
                    .doc()
                    .delete_range(&EditCtx::local("", ""), StoryRange::new("body", 3, 4))
                    .map(|_| ()),
            }
            .unwrap();
            let epoch = engine.stats().frame_epoch;
            engine.apply_and_layout("body", epoch).unwrap();
            assert_cold(&request);
        }
        request["renderEnv"]["revisionPreview"] = json!({"10": decision});
        engine
            .layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        let epoch = engine.stats().frame_epoch;
        engine
            .build_display_list_frame(&extras(&request.to_string()), epoch)
            .unwrap();
        assert_cold(&request);
    }
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

    fresh
        .doc()
        .set_note_separator_state(engine.doc().note_separator_state().unwrap());
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

    fresh
        .doc()
        .set_note_separator_state(engine.doc().note_separator_state().unwrap());
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
fn multiple_layouts_before_a_display_build_preserve_body_and_note_damage() {
    let body: String = (1..=10)
        .map(|page| {
            let mut content = if page == 1 {
                String::new()
            } else {
                "<w:pPr><w:pageBreakBefore/></w:pPr>".to_owned()
            };
            content += &fixture::r(&format!("Page {page}"));
            if page == 4 {
                content += r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#;
            }
            fixture::p(&format!("{:08X}", 0x7000_0000 + page), &content)
        })
        .collect();
    let bytes = fixture::with_body_and_note(
        &small_page(&body),
        &fixture::p("70000100", &fixture::r("Note")),
    );
    let (engine, request) = bulk_laid_out(&bytes, 9321);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    let initial = display(&engine);
    assert_eq!(initial["pages"].as_array().unwrap().len(), 10);
    let anchor = initial["pages"][3]["noteAreas"][0]["notes"][0]["anchorDocStart"]
        .as_i64()
        .unwrap();
    let before = engine.stats();

    bulk_edit(
        &engine,
        json!({
            "op": "insertText", "at": "end", "text": "x",
            "target": {"kind": "paragraph", "story": "body", "paraId": "70000002"},
        }),
    );
    let early = engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(pages(&early), 10);
    assert_eq!(engine.stats().display_builds, before.display_builds);
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before.incremental_pagination_calls + 1
    );

    bulk_edit(
        &engine,
        json!({
            "op": "insertText", "at": "end", "text": "y",
            "target": {"kind": "paragraph", "story": "body", "paraId": "70000008"},
        }),
    );
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(pages(&layout), 10);
    assert_eq!(engine.stats().display_builds, before.display_builds);
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before.incremental_pagination_calls + 2
    );
    engine
        .build_display_list_frame(&extras(&request), before.frame_epoch)
        .unwrap();
    assert_eq!(
        engine.stats().incremental_display_builds,
        before.incremental_display_builds + 1
    );
    let actual = display(&engine);
    let (expected_layout, expected) = fresh(&engine, &request, 9322);
    assert_eq!(layout, expected_layout);
    assert_eq!(
        expected["pages"][3]["noteAreas"][0]["notes"][0]["anchorDocStart"],
        anchor + 1
    );
    assert_eq!(
        actual["pages"][3]["noteAreas"],
        expected["pages"][3]["noteAreas"]
    );
    assert_eq!(actual["pages"][1], expected["pages"][1]);
    assert_eq!(
        serde_json::to_vec(&actual).unwrap(),
        serde_json::to_vec(&expected).unwrap()
    );
}

#[test]
fn an_intermediate_page_build_before_undo_preserves_display_positions() {
    let body: String = (1..=7)
        .map(|page| {
            let mut content = if page == 1 {
                String::new()
            } else {
                "<w:pPr><w:pageBreakBefore/></w:pPr>".to_owned()
            };
            content += &fixture::r(&format!("Page {page}"));
            if page == 4 {
                content += r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#;
            }
            fixture::p(&format!("{:08X}", 0x7100_0000 + page), &content)
        })
        .collect();
    let bytes = fixture::with_body_and_note(
        &small_page(&body),
        &fixture::p("71000100", &fixture::r("Note")),
    );
    let (engine, request) = bulk_laid_out(&bytes, 9323);
    engine.set_display_window(Some(0..6));
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    let initial = display(&engine);
    assert_eq!(initial["pages"].as_array().unwrap().len(), 7);
    assert_eq!(initial["pages"][6]["unbuilt"], true);
    let edits = docx_edit::EditingDoc::new(9328);
    edits
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let undo = UndoSession::new();
    undo.track(&edits);
    let before = engine.stats();

    edits
        .insert_text(
            &EditCtx::local("", ""),
            edits.paragraph_mark_position("71000002").unwrap(),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    engine
        .doc()
        .apply_host_update_v1(&edits.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        pages(&engine.layout_document_with_regions_json(&request).unwrap()),
        7
    );
    engine
        .build_display_pages_frame(&[6], before.frame_epoch)
        .unwrap();
    let intermediate = display(&engine);
    assert_ne!(intermediate["pages"][6]["unbuilt"], true);

    assert!(undo.undo());
    engine
        .doc()
        .apply_host_update_v1(&edits.encode_state_as_update_v1())
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(pages(&layout), 7);
    assert_eq!(engine.stats().display_builds, before.display_builds);
    assert_eq!(engine.stats().pagination_calls, before.pagination_calls + 2);
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before.incremental_pagination_calls + 2
    );
    let frame = engine
        .build_display_list_frame(&extras(&request), engine.stats().frame_epoch)
        .unwrap();
    let operations = u32::from_le_bytes(frame[52..56].try_into().unwrap()) as usize;
    assert!((0..operations).any(|operation| {
        let index = FRAME_HEADER_LEN + operation * PAGE_OP_LEN + 4;
        u32::from_le_bytes(frame[index..index + 4].try_into().unwrap()) == 6
    }));
    assert_eq!(
        engine.stats().incremental_display_builds,
        before.incremental_display_builds + 1
    );
    let actual = display(&engine);
    let (expected_layout, expected) = fresh(&engine, &request, 9324);
    assert_eq!(layout, expected_layout);
    assert_ne!(intermediate["pages"][6], expected["pages"][6]);
    assert_eq!(
        serde_json::to_vec(&actual).unwrap(),
        serde_json::to_vec(&expected).unwrap()
    );
}

#[test]
fn accumulated_display_damage_is_bounded_without_changing_single_layout_builds() {
    let page_count = 258;
    let body: String = (1..=page_count)
        .map(|page| {
            let content = format!(
                "{}{}",
                if page == 1 {
                    ""
                } else {
                    "<w:pPr><w:pageBreakBefore/></w:pPr>"
                },
                fixture::r(&format!("Page {page} A"))
            );
            fixture::p(&format!("{:08X}", 0x7200_0000 + page), &content)
        })
        .collect();
    let (engine, request) = bulk_laid_out(&fixture::with_body(&small_page(&body)), 9325);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    assert_eq!(
        display(&engine)["pages"].as_array().unwrap().len(),
        page_count
    );
    let replace = |page, text| {
        let para_id = format!("{:08X}", 0x7200_0000 + page);
        let offset = format!("Page {page} ").len();
        bulk_edit(
            &engine,
            json!({
                "op": "replaceText", "text": text,
                "target": {
                    "kind": "range", "story": "body", "view": "accepted",
                    "start": {"paraId": para_id, "offset": offset},
                    "end": {"paraId": para_id, "offset": offset + 1},
                },
            }),
        );
    };
    let before = engine.stats();
    for page in 1..=page_count {
        replace(page, "B");
    }
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(pages(&layout), page_count);
    engine
        .build_display_list_frame(&extras(&request), before.frame_epoch)
        .unwrap();
    assert_eq!(
        engine.stats().incremental_display_builds,
        before.incremental_display_builds + 1
    );
    assert_eq!(
        engine.stats().rebuilt_display_pages - before.rebuilt_display_pages,
        page_count as u64
    );
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9326));

    let before = engine.stats();
    let mut layout = String::new();
    for page in 1..=page_count {
        replace(page, "C");
        layout = engine.layout_document_with_regions_json(&request).unwrap();
        assert_eq!(pages(&layout), page_count);
    }
    assert_eq!(engine.stats().display_builds, before.display_builds);
    assert_eq!(
        engine.stats().pagination_calls,
        before.pagination_calls + page_count as u64
    );
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before.incremental_pagination_calls + page_count as u64
    );
    engine
        .build_display_list_frame(&extras(&request), before.frame_epoch)
        .unwrap();
    assert_eq!(
        engine.stats().incremental_display_builds,
        before.incremental_display_builds
    );
    assert_eq!(
        engine.stats().rebuilt_display_pages - before.rebuilt_display_pages,
        page_count as u64
    );
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9327));
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

fn pages(layout: &str) -> usize {
    serde_json::from_str::<Value>(layout).unwrap()["layout"]["pages"]
        .as_array()
        .unwrap()
        .len()
}

fn small_page(body: &str) -> String {
    format!(
        concat!(
            "{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" ",
            "w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" ",
            "w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>"
        ),
        body
    )
}

#[test]
fn shrinking_a_kept_follower_pulls_its_keep_with_next_run_back() {
    let filler = "filler words to fill most of the first page ".repeat(20);
    let follower = "a follower kept on its lines with its keep with next head ".repeat(6);
    let body = small_page(&format!(
        "{}{}{}",
        fixture::p("50000001", &fixture::r(&filler)),
        fixture::p(
            "50000002",
            &format!("<w:pPr><w:keepNext/></w:pPr>{}", fixture::r("Kept heading"))
        ),
        fixture::p(
            "50000003",
            &format!("<w:pPr><w:keepLines/></w:pPr>{}", fixture::r(&follower))
        ),
    ));
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9314);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    assert_eq!(
        pages(&engine.layout_document_with_regions_json(&request).unwrap()),
        2
    );
    engine
        .build_display_list_frame(&extras(&request), 1)
        .unwrap();

    let start = engine.doc().paragraphs("body").unwrap()[..2]
        .iter()
        .map(|paragraph| paragraph.text.encode_utf16().count() as u32 + 1)
        .sum::<u32>();
    engine
        .doc()
        .delete_range(
            &EditCtx::local("", ""),
            StoryRange::new("body", start + 1, start + follower.len() as u32),
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    engine
        .build_display_list_frame(&extras(&request), 2)
        .unwrap();
    assert_eq!(pages(&layout), 1);
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9315));
}

/// Pages of text after a paragraph whose endnote is laid out on the last page.
fn endnote_after_pages(client_id: u64) -> (EngineSession, String) {
    let filler = (0..40)
        .map(|index| {
            fixture::p(
                &format!("{:08X}", 0x6000_0010 + index),
                &fixture::r(&"enough words to spread over several pages ".repeat(4)),
            )
        })
        .collect::<String>();
    let body = small_page(&format!(
        "{}{filler}",
        fixture::p(
            "60000001",
            &format!(
                r#"{}<w:r><w:endnoteReference w:id="1"/></w:r>{}"#,
                fixture::r("Alpha"),
                fixture::r("beta")
            )
        )
    ));
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), client_id);
    engine
        .build_display_list_frame(&extras(&request), 0)
        .unwrap();
    assert!(pages(&engine.layout_document_with_regions_json(&request).unwrap()) > 2);
    engine
        .build_display_list_frame(&extras(&request), 1)
        .unwrap();
    (engine, request)
}

#[test]
fn a_same_length_edit_around_an_endnote_reference_moves_its_backlink() {
    let (engine, request) = endnote_after_pages(9316);
    let ctx = EditCtx::local("", "");
    engine
        .doc()
        .insert_text(&ctx, Position::new("body", 0), "x", FormatPolicy::Inherit)
        .unwrap();
    engine
        .doc()
        .delete_range(&ctx, StoryRange::new("body", 7, 8))
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    engine
        .build_display_list_frame(&extras(&request), 2)
        .unwrap();
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9317));
}

#[test]
fn a_body_edit_batched_with_an_endnote_edit_redraws_the_endnote() {
    let (engine, request) = endnote_after_pages(9318);
    let ctx = EditCtx::local("", "");
    engine
        .doc()
        .insert_text(&ctx, Position::new("body", 0), "x", FormatPolicy::Inherit)
        .unwrap();
    engine
        .doc()
        .delete_range(&ctx, StoryRange::new("body", 1, 2))
        .unwrap();
    engine
        .doc()
        .insert_text(&ctx, Position::new("en:1", 1), "y", FormatPolicy::Inherit)
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    engine
        .build_display_list_frame(&extras(&request), 2)
        .unwrap();
    assert_eq!((layout, display(&engine)), fresh(&engine, &request, 9319));
}

#[test]
fn a_page_side_float_leaving_the_text_area_remeasures_its_neighbours() {
    let shape = concat!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" "#,
        r#"simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" "#,
        r#"allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page">"#,
        r#"<wp:posOffset>2857500</wp:posOffset></wp:positionH><wp:positionV "#,
        r#"relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
        r#"<wp:extent cx="914400" cy="914400"/><wp:wrapSquare wrapText="bothSides"/>"#,
        r#"<wp:docPr id="1" name="Page side shape"/><a:graphic><a:graphicData "#,
        r#"uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>"#,
        r#"<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" "#,
        r#"cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>"#,
        r#"<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    );
    let words = "words that wrap beside a shape on the page side ".repeat(5);
    let margins = |left: u32, right: u32| {
        format!(
            concat!(
                "<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"8640\"/><w:pgMar w:top=\"720\" ",
                "w:right=\"{}\" w:bottom=\"720\" w:left=\"{}\" w:header=\"300\" ",
                "w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>"
            ),
            right, left
        )
    };
    let body = |left: u32, right: u32| {
        format!(
            "{}{}{}",
            fixture::p("70000001", &format!("{shape}{}", fixture::r(&words))),
            fixture::p("70000002", &fixture::r(&words)),
            margins(left, right),
        )
    };
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body(2880, 720)), 9318);
    engine.layout_document_with_regions_json(&request).unwrap();
    let swapped = fixture::region_request(
        &engine,
        &fixture::with_body(&body(720, 2880)),
        serde_json::from_str::<Value>(&request).unwrap()["measurement"]["fontChains"]
            .as_object()
            .unwrap()
            .values()
            .next()
            .unwrap()[0]
            .as_u64()
            .unwrap() as u32,
    )
    .to_string();
    let layout = engine.layout_document_with_regions_json(&swapped).unwrap();
    engine
        .build_display_list_frame(&extras(&swapped), 0)
        .unwrap();
    assert_ne!(
        layout,
        fresh(&engine, &request, 9319).0,
        "the swap moves the shape out"
    );
    assert_eq!((layout, display(&engine)), fresh(&engine, &swapped, 9320));
}
