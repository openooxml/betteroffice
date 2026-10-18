#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::structured::{PageDiagnosticCode, PageExportOptions, RevisionView};
use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, Position, SegmentContent, bridge::RenderEnv,
    seed_from_docx,
};
use docx_layout::types::{BlockId, LayoutBlock};
use serde_json::{Value, json};

const PAGE_BREAK: &str = r#"<w:r><w:br w:type="page"/></w:r>"#;

fn body_document(body: &str) -> Vec<u8> {
    let bytes = fixture::with_body(&format!("{body}<w:sectPr/>"));
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
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

fn page_texts(body: &str) -> Vec<String> {
    let bytes = body_document(body);
    let (engine, _) = fixture::laid_out(&bytes, 9403);
    let layout: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
    layout_page_texts(&layout)
}

fn layout_page_texts(layout: &Value) -> Vec<String> {
    layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| fragment["resolvedLines"].as_array().into_iter().flatten())
                .flat_map(|line| line["segments"].as_array().into_iter().flatten())
                .filter_map(|segment| segment["text"].as_str())
                .collect()
        })
        .collect()
}

#[test]
fn a_terminal_page_break_in_the_last_text_paragraph_keeps_two_pages() {
    let body = fixture::p("71000000", &format!("{}{PAGE_BREAK}", fixture::r("A")));
    assert_eq!(page_texts(&body), ["A", ""]);
}

#[test]
fn a_page_break_between_text_in_the_last_paragraph_keeps_two_pages() {
    let body = fixture::p(
        "71000000",
        &format!("{}{PAGE_BREAK}{}", fixture::r("A"), fixture::r("B")),
    );
    assert_eq!(page_texts(&body).len(), 2);
}

#[test]
fn a_terminal_break_only_paragraph_does_not_add_a_page() {
    let body = format!(
        "{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
    );
    assert_eq!(page_texts(&body), ["A"]);
}

#[test]
fn a_page_break_before_more_text_opens_a_page() {
    let body = format!(
        "{}{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
        fixture::p("71000002", &fixture::r("B")),
    );
    assert_eq!(page_texts(&body), ["A", "B"]);
}

#[test]
fn consecutive_break_only_paragraphs_preserve_an_authored_empty_page() {
    let body = format!(
        "{}{}{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
        fixture::p("71000002", PAGE_BREAK),
        fixture::p("71000003", &fixture::r("B")),
    );
    assert_eq!(page_texts(&body), ["A", "", "B"]);
}

#[test]
fn an_empty_paragraph_after_a_page_break_consumes_the_new_page() {
    for body in [
        fixture::p("71000000", &format!("{}{PAGE_BREAK}", fixture::r("A"))),
        format!(
            "{}{}",
            fixture::p("71000000", &fixture::r("A")),
            fixture::p("71000001", PAGE_BREAK),
        ),
    ] {
        let body = format!("{body}{}", fixture::p("71000002", ""));
        assert_eq!(page_texts(&body), ["A", ""]);
    }
}

#[test]
fn content_after_a_leading_break_in_the_last_paragraph_consumes_the_new_page() {
    for (tail, expected) in [
        (fixture::r("B"), "B"),
        (r#"<w:r><w:br/></w:r>"#.to_owned(), ""),
    ] {
        let body = format!(
            "{}{}",
            fixture::p("71000000", &fixture::r("A")),
            fixture::p("71000001", &format!("{PAGE_BREAK}{tail}")),
        );
        assert_eq!(page_texts(&body), ["A", expected]);
    }
}

#[test]
fn retiring_a_terminal_break_preserves_earlier_authored_empty_pages() {
    let body = format!(
        "{}{}{}{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
        fixture::p("71000002", PAGE_BREAK),
        fixture::p("71000003", &fixture::r("B")),
        fixture::p("71000004", PAGE_BREAK),
    );
    assert_eq!(page_texts(&body), ["A", "", "B"]);
}

#[test]
fn consecutive_terminal_break_only_paragraphs_drop_only_the_last_break() {
    let body = format!(
        "{}{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
        fixture::p("71000002", PAGE_BREAK),
    );
    assert_eq!(page_texts(&body), ["A", ""]);
}

#[test]
fn source_content_after_a_break_only_paragraph_keeps_the_break() {
    let body = format!(
        r#"{}{}<w:bookmarkStart w:id="4" w:name="tail"/><w:bookmarkEnd w:id="4"/>"#,
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
    );
    assert_eq!(page_texts(&body), ["A", ""]);
}

fn assert_endnotes_keep_terminal_break(position: &str) {
    let body = format!(
        r#"{}{}<w:sectPr><w:endnotePr><w:pos w:val="{position}"/></w:endnotePr></w:sectPr>"#,
        fixture::p(
            "71000000",
            &format!(
                r#"{}<w:r><w:endnoteReference w:id="1"/></w:r>"#,
                fixture::r("A")
            ),
        ),
        fixture::p("71000001", PAGE_BREAK),
    );
    let (engine, _) = fixture::laid_out(&fixture::with_body(&body), 9404);
    let layout: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
    assert_eq!(layout_page_texts(&layout).len(), 2);
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        blocks.as_array().unwrap().last().unwrap()["kind"],
        "pageBreak"
    );
}

#[test]
fn document_end_endnotes_keep_a_terminal_break_only_paragraphs_break() {
    assert_endnotes_keep_terminal_break("docEnd");
}

#[test]
fn final_section_endnotes_keep_a_terminal_break_only_paragraphs_break() {
    assert_endnotes_keep_terminal_break("sectEnd");
}

#[test]
fn separator_only_endnotes_do_not_keep_a_terminal_break_only_paragraphs_break() {
    let body = format!(
        "{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
    );
    let mut parts = ooxml_opc::unzip_parts(&body_document(&body)).unwrap();
    let separators: String = [
        (-1, "separator"),
        (0, "continuationSeparator"),
        (2, "continuationNotice"),
    ]
    .into_iter()
    .map(|(id, kind)| format!(r#"<w:endnote w:id="{id}" w:type="{kind}"><w:p/></w:endnote>"#))
    .collect();
    let (_, endnotes) = parts
        .iter_mut()
        .find(|(name, _)| name == "word/endnotes.xml")
        .unwrap();
    *endnotes = format!(
        r#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{separators}</w:endnotes>"#
    )
    .into_bytes();
    let (engine, _) = fixture::laid_out(&ooxml_opc::rezip_parts(&parts).unwrap(), 9412);
    let layout: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
    assert_eq!(layout_page_texts(&layout), ["A"]);
}

#[test]
fn terminal_break_omission_respects_final_section_footnote_placement() {
    for (position, pages) in [("docEnd", 2), ("sectEnd", 2), ("pageBottom", 1)] {
        let body = format!(
            r#"{}{}<w:sectPr><w:footnotePr><w:pos w:val="{position}"/></w:footnotePr></w:sectPr>"#,
            fixture::p(
                "71000000",
                &format!(
                    r#"{}<w:r><w:footnoteReference w:id="1"/></w:r>"#,
                    fixture::r("A")
                ),
            ),
            fixture::p("71000001", PAGE_BREAK),
        );
        let mut parts = ooxml_opc::unzip_parts(&fixture::with_body(&body)).unwrap();
        let (_, endnotes) = parts
            .iter_mut()
            .find(|(name, _)| name == "word/endnotes.xml")
            .unwrap();
        *endnotes = br#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#.to_vec();
        let (engine, _) = fixture::laid_out(&ooxml_opc::rezip_parts(&parts).unwrap(), 9413);
        let layout: Value = serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
        assert_eq!(layout_page_texts(&layout).len(), pages, "{position}");
    }
}

#[test]
fn a_page_prefix_ending_on_a_mid_document_break_still_opens_its_last_page() {
    let mut body = String::new();
    for index in 0..10 {
        body.push_str(&fixture::p(
            &format!("{:08X}", 0x7100_0000 + index * 2),
            &fixture::r(&format!("Page {index}")),
        ));
        body.push_str(&fixture::p(
            &format!("{:08X}", 0x7100_0001 + index * 2),
            PAGE_BREAK,
        ));
    }
    body.push_str(&fixture::p("72000000", PAGE_BREAK));
    body.push_str(&fixture::p("72000001", &fixture::r("After probe")));
    for index in 0..10 {
        body.push_str(&fixture::p(
            &format!("{:08X}", 0x7300_0000 + index * 2),
            PAGE_BREAK,
        ));
        body.push_str(&fixture::p(
            &format!("{:08X}", 0x7300_0001 + index * 2),
            &fixture::r("Later text"),
        ));
    }
    let bytes = body_document(&body);
    let (full, request) = fixture::laid_out(&bytes, 9405);
    let full: Value = serde_json::from_str(&full.retained_layout_json().unwrap()).unwrap();
    let prefix = EngineSession::new(9406);
    seed_from_docx(prefix.doc(), &bytes).unwrap();
    let prefix: Value = serde_json::from_str(
        &prefix
            .layout_document_with_regions_prefix_retained_json(&request, 9)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(prefix["provisional"], true);
    let prefix_text = layout_page_texts(&prefix);
    let full_text = layout_page_texts(&full);
    assert_eq!(prefix_text.len(), 12);
    assert_eq!(prefix_text.last().unwrap(), "");
    assert_eq!(prefix_text[..11], full_text[..11]);
    assert_eq!(full_text[11], "After probe");
}

#[test]
fn omitting_a_terminal_break_preserves_the_story_and_paragraph_mapping() {
    let body = format!(
        "{}{}",
        fixture::p("71000000", &fixture::r("A")),
        fixture::p("71000001", PAGE_BREAK),
    );
    let bytes = body_document(&body);
    let engine = EngineSession::new(9407);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let before = engine.doc().story_segments("body").unwrap();
    assert!(matches!(
        &before.last().unwrap().content,
        SegmentContent::OtherEmbed { kind, .. } if kind == "pageBreak"
    ));
    let (blocks, map) = docx_edit::bridge::yrs_doc_to_mapped_layout_blocks(
        engine.doc(),
        "body",
        &RenderEnv::default(),
    )
    .unwrap();
    assert_eq!(blocks.len(), 2);
    assert_eq!(map.paragraphs.len(), 2);
    let LayoutBlock::Paragraph(paragraph) = &blocks[1] else {
        panic!("expected the final paragraph");
    };
    assert_eq!(paragraph.id, BlockId::Str(map.paragraphs[1].1.clone()));
    assert_eq!(
        (paragraph.pm_start, paragraph.pm_end),
        (Some(3.0), Some(5.0))
    );
    assert_eq!(map.paragraph_at(3), Some(1));
    assert_eq!(engine.doc().story_segments("body").unwrap(), before);

    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let request = fixture::region_request(&engine, &bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let exported = engine
        .export_structured_with_pages(&PageExportOptions::new(RevisionView::Markup))
        .unwrap()
        .content;
    assert!(exported.layout.diagnostics.iter().all(|diagnostic| {
        diagnostic.code != PageDiagnosticCode::NotLaidOut
            && diagnostic.code != PageDiagnosticCode::UnmappedContent
    }));
    assert_eq!(engine.doc().story_segments("body").unwrap(), before);

    let plan = engine.doc().paragraph_save_plan();
    let wire = docx_parse::parse_docx_s9_wire(&bytes, Default::default()).unwrap();
    let projection = docx_parse::parse_docx_s8_projection(&bytes).unwrap();
    let request = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": projection.body,
        "relationshipEntries": wire.document.package.relationship_entries,
        "headerEntries": wire.document.package.header_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {
            "assignments": plan.assignments.iter().map(|(part, ordinal, para_id)| {
                json!({"part": part, "ordinal": ordinal, "paraId": para_id})
            }).collect::<Vec<_>>(),
            "patchedParts": plan.patched_parts.iter().map(|(part, para_ids)| {
                json!({"part": part, "paraIds": para_ids})
            }).collect::<Vec<_>>(),
        },
    }))
    .unwrap();
    let saved = docx_parse::serializer::s13::write_docx_s13(request, &bytes).unwrap();
    let original_parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let main_part = |parts: Vec<(String, Vec<u8>)>| {
        parts
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1
    };
    assert_eq!(main_part(saved_parts), main_part(original_parts));
}

#[test]
fn an_edit_before_page_breaks_repaginates_only_its_own_page() {
    let paragraph = |id: u32| fixture::p(&format!("{id:08X}"), &fixture::r("text on its own page"));
    let page_break =
        |id: u32| fixture::p(&format!("{id:08X}"), r#"<w:r><w:br w:type="page"/></w:r>"#);
    let body: String = (0..6_u32)
        .map(|index| {
            format!(
                "{}{}",
                paragraph(0x7000_0000 + index * 2),
                page_break(0x7000_0001 + index * 2)
            )
        })
        .collect();
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9401);
    let before = engine.stats();
    assert!(before.retained_pages >= 6);

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 0),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );
    assert_eq!(after.rebuilt_pages, 1, "later pages converge");

    let fresh = EngineSession::new(9402);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        layout
    );
}
