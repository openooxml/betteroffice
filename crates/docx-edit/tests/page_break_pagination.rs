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

fn section(kind: &str) -> String {
    format!(
        r#"<w:sectPr>{kind}<w:pgSz w:w="7200" w:h="4320"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#
    )
}

fn render_body(body: &str) -> Value {
    let bytes = fixture::with_body(body);
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut sections: Vec<_> = package
        .document
        .sections
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    sections.push(json!({"properties": package.document.final_section_properties}));
    let engine = EngineSession::new(9410);
    docx_edit::seed_from_docx(engine.doc(), &bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": sections, "settings": package.settings},
                    "measurement": {
                        "fontChains": {"arial|0|0": [font]},
                        "defaults": {"fontFamily": "Arial", "fontSize": 12}
                    }
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn text_page(output: &Value, text: &str) -> u64 {
    let block = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["runs"]
                .as_array()
                .is_some_and(|runs| runs.iter().any(|run| run["text"] == text))
        })
        .unwrap();
    fragment_page(output, &block["block"]["id"]).0
}

fn fragment_page<'a>(output: &'a Value, block_id: &Value) -> (u64, &'a Value) {
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .find_map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .find(|fragment| fragment["blockId"] == *block_id)
                .map(|fragment| (page["number"].as_u64().unwrap(), fragment))
        })
        .unwrap()
}

fn assert_section_mark(output: &Value, page: u64) {
    let mark = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["kind"] == "paragraph"
                && measured["measure"]["lines"]
                    .as_array()
                    .is_some_and(Vec::is_empty)
        })
        .unwrap();
    let (number, fragment) = fragment_page(output, &mark["block"]["id"]);
    assert_eq!(number, page);
    assert_eq!(fragment["height"], 0.0);
    assert_eq!(fragment["pmStart"], mark["block"]["pmStart"]);
    assert_eq!(fragment["pmEnd"], mark["block"]["pmEnd"]);
    assert!(fragment["pmStart"].is_number());
    assert_eq!(fragment["fromLine"], 0);
    assert_eq!(fragment["toLine"], 0);
}

fn assert_two_pages(output: &Value) {
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(output, "A"), 1);
    assert_eq!(text_page(output, "B"), 2);
    assert_eq!(output["layout"]["pages"][1]["sectionIndex"], 1);
}

#[test]
fn bare_section_mark_after_break_only_paragraph_reuses_the_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 2);
}

#[test]
fn bare_section_mark_after_trailing_page_break_reuses_the_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 2);
}

#[test]
fn page_break_before_section_mark_keeps_a_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_eq!(output["layout"]["pages"][2]["sectionIndex"], 1);
    assert_section_mark(&output, 2);
}

#[test]
fn bare_section_mark_does_not_flush_spacing_past_the_body_bottom() {
    let geometry = section("");
    for (line, after) in [(2880, 0), (2640, 1080)] {
        let output = render_body(&format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="exact"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="240" w:after="240"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
        ));
        assert_two_pages(&output);
        assert_section_mark(&output, 1);
    }
}

#[test]
fn bare_section_mark_after_a_full_page_table_does_not_add_a_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2880" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:before="240"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| measured["block"]["kind"] == "table")
        .unwrap();
    assert_eq!(fragment_page(&output, &table["block"]["id"]).0, 2);
    assert_section_mark(&output, 2);
}

#[test]
fn keep_next_does_not_reserve_spacing_for_a_bare_section_mark() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>PREFIX</w:t></w:r></w:p><w:p><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="2400" w:lineRule="exact"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="480"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 1);
}

#[test]
fn authored_empty_paragraph_after_manual_break_keeps_its_line() {
    for separate_break in [false, true] {
        let first = if separate_break {
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>"#
        } else {
            r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p>"#
        };
        let output = render_body(&format!(
            r#"{first}<w:p/><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section("")
        ));
        assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), 2);
        let measured = output["measured"].as_array().unwrap();
        let empty = measured
            .iter()
            .rev()
            .find(|block| {
                block["block"]["kind"] == "paragraph"
                    && block["block"]["runs"].as_array().is_some_and(Vec::is_empty)
            })
            .unwrap();
        let (page, fragment) = fragment_page(&output, &empty["block"]["id"]);
        assert_eq!(page, 2);
        assert!(fragment["height"].as_f64().unwrap() > 0.0);
    }
}

#[test]
fn authored_empty_paragraphs_before_section_mark_keep_a_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p/><w:p/><w:p/><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    let empties: Vec<_> = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|measured| {
            measured["block"]["kind"] == "paragraph"
                && measured["block"]["runs"]
                    .as_array()
                    .is_some_and(Vec::is_empty)
                && measured["measure"]["lines"]
                    .as_array()
                    .is_some_and(|lines| !lines.is_empty())
        })
        .collect();
    assert_eq!(empties.len(), 3);
    for empty in empties {
        let (page, fragment) = fragment_page(&output, &empty["block"]["id"]);
        assert_eq!(page, 2);
        assert!(fragment["height"].as_f64().unwrap() > 0.0);
    }
    assert_section_mark(&output, 2);
}

#[test]
fn odd_page_section_after_bare_mark_keeps_the_parity_filler() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(""),
        section(r#"<w:type w:val="oddPage"/>"#)
    ));
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_eq!(pages[1]["parityFiller"], true);
    assert_section_mark(&output, 1);
}

#[test]
fn manual_break_before_bare_mark_leaves_parity_fillers_blank() {
    for (kind, preceding, expected_page) in [
        ("oddPage", "", 3),
        (
            "evenPage",
            r#"<w:p><w:r><w:t>C</w:t><w:br w:type="page"/></w:r></w:p>"#,
            4,
        ),
    ] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p>{preceding}<w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section(""),
            section(&format!(r#"<w:type w:val="{kind}"/>"#))
        ));
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), expected_page as usize);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), expected_page);
        let filler = &pages[expected_page as usize - 2];
        assert_eq!(filler["parityFiller"], true);
        assert!(
            filler["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .all(|fragment| fragment["height"] == 0.0)
        );
        assert_section_mark(&output, expected_page - 1);
    }
}

#[test]
fn authored_empty_paragraph_between_manual_breaks_keeps_a_blank_page() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p/><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
}

#[test]
fn section_mark_before_a_continuous_section_keeps_spacing() {
    for resized in [false, true] {
        let mut next = section(r#"<w:type w:val="continuous"/>"#);
        if resized {
            next = next.replace(r#"w:w="7200""#, r#"w:w="8640""#);
        }
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="4320" w:after="4320"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{next}"#,
            section("")
        ));
        assert!(output["layout"]["pages"].as_array().unwrap().len() >= 2);
        assert_eq!(text_page(&output, "A"), 1);
        assert!(text_page(&output, "B") >= 2);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn section_mark_spacing_participates_in_column_balancing() {
    let columns = r#"<w:cols w:num="2" w:space="360"/>"#;
    let mut positions = Vec::new();
    for spacing in [0, 4320] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="{spacing}" w:after="{spacing}"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section(columns),
            section(&format!(r#"<w:type w:val="continuous"/>{columns}"#))
        ));
        if spacing == 0 {
            assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
            assert_eq!(text_page(&output, "B"), 1);
            assert_section_mark(&output, 1);
        }
        positions.push(
            output["layout"]["pages"][0]["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .map(|fragment| (fragment["x"].clone(), fragment["y"].clone()))
                .collect::<Vec<_>>(),
        );
    }
    assert_ne!(positions[0], positions[1]);
}

#[test]
fn anchored_cover_section_mark_keeps_before_spacing() {
    let geometry = r#"<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="500" w:right="500" w:bottom="500" w:left="500"/></w:sectPr>"#;
    let output = render_body(&format!(
        r#"<w:p/><w:p/><w:p><w:pPr><w:spacing w:before="13000" w:after="200"/>{geometry}</w:pPr>{}</w:p><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1800" w:lineRule="exact"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#,
        anchored_text_box()
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(&output, "B"), 2);
    let anchor = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|mb| {
            mb["block"]["attrs"]["hasSourceContent"] == true
                && mb["block"]["runs"].as_array().is_some_and(Vec::is_empty)
        })
        .unwrap();
    assert!(
        (anchor["block"]["attrs"]["spacing"]["before"]
            .as_f64()
            .unwrap()
            - 13000.0 / 15.0)
            .abs()
            < 1e-6
    );
    assert_eq!(fragment_page(&output, &anchor["block"]["id"]).0, 1);
    assert!(output["measured"].as_array().unwrap().iter().any(|mb| {
        mb["block"]["kind"] == "shape"
            && mb["block"]["innerText"]
                .as_array()
                .is_some_and(|text| !text.is_empty())
    }));
}

fn anchored_text_box() -> &'static str {
    concat!(
        r#"<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">"#,
        r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>"#,
        r#"<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="Cover"/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>"#,
        r#"<wps:txbx><w:txbxContent><w:p><w:r><w:t>COVER</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
    )
}

#[test]
fn anchored_section_mark_before_next_page_keeps_spacing() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="13000"/>{geometry}</w:pPr>{}</w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#,
        anchored_text_box()
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
}

#[test]
fn two_column_continuous_section_mark_after_page_break_keeps_blank_page() {
    let ending = section(r#"<w:type w:val="continuous"/><w:cols w:num="2" w:space="360"/>"#);
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{ending}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn section_mark_after_whitespace_and_page_break_keeps_blank_page() {
    let geometry = section("");
    for whitespace in [
        "<w:tab/>".repeat(7),
        r#"<w:t xml:space="preserve">   </w:t>"#.to_owned(),
    ] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r>{whitespace}<w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
        ));
        assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), 3);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn section_mark_after_page_break_with_trailing_space_keeps_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn continuous_section_mark_after_full_page_table_keeps_blank_page() {
    let output = render_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2880" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:before="240"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(r#"<w:type w:val="continuous"/>"#),
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn one_column_continuous_section_mark_after_terminal_break_reuses_page() {
    for resized in [false, true] {
        let incoming = if resized {
            section("").replace(r#"w:w="7200""#, r#"w:w="8640""#)
        } else {
            section("")
        };
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/>{}</w:pPr><w:bookmarkStart w:id="0" w:name="boundary"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{incoming}"#,
            section(r#"<w:type w:val="continuous"/>"#),
        ));
        assert_two_pages(&output);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn framed_section_mark_after_page_break_keeps_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:framePr w:w="1440" w:hAnchor="margin" w:vAnchor="text"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn section_mark_entering_two_columns_after_page_break_keeps_blank_page() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(""),
        section(r#"<w:cols w:num="2" w:space="360"/>"#)
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn genuinely_empty_section_keeps_its_line_and_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(&output, "B"), 2);
    assert!(
        output["layout"]["pages"][0]["fragments"][0]["height"]
            .as_f64()
            .unwrap()
            > 0.0
    );
}

fn staggered_row(exact: bool) -> Vec<u8> {
    let cells: String = [("left", 90, 240), ("right", 85, 260)]
        .into_iter()
        .enumerate()
        .map(|(cell, (prefix, count, spacing))| {
            let paragraphs: String = (0..count)
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7100_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("{prefix}{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let height = if exact {
        r#"<w:trPr><w:trHeight w:val="7200" w:hRule="exact"/></w:trPr>"#
    } else {
        ""
    };
    fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>{height}{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ))
}

fn staggered_row_layout(exact: bool) -> (Value, Value) {
    let (engine, request) = fixture::laid_out(&staggered_row(exact), 9403);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    (
        serde_json::from_str(&output).unwrap(),
        serde_json::from_str(&display).unwrap(),
    )
}

#[test]
fn centred_cells_with_staggered_lines_split_within_the_page_body() {
    let (output, display) = staggered_row_layout(false);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let mut fragment_count = 0;
    let mut texts = [String::new(), String::new()];
    for (page, painted) in pages.iter().zip(display["pages"].as_array().unwrap()) {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        if fragments.is_empty() {
            continue;
        }
        assert_eq!(fragments.len(), 1);
        fragment_count += 1;
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        let clips = fragment["cellClips"].as_array().unwrap();
        assert_eq!(clips.len(), 2);
        assert_ne!(clips[0]["bottom"], clips[1]["bottom"]);
        for row in fragment["rowStart"].as_u64().unwrap()..fragment["rowEnd"].as_u64().unwrap() {
            assert!(
                painted["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|primitive| {
                        primitive["cell"]["row"].as_u64() == Some(row)
                            && primitive["text"]
                                .as_str()
                                .is_some_and(|text| !text.is_empty())
                    }),
                "row {row} has no line text in {fragment}"
            );
        }
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            let measured_cell = &output["measured"][0]["measure"]["rows"][0]["cells"][cell];
            let line = &measured_cell["blocks"][0]["lines"][0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_bottom = line_top + line["lineHeight"].as_f64().unwrap();
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
    }
    assert!(fragment_count > 1);
    for (cell, (prefix, count)) in [("left", 90), ("right", 85)].into_iter().enumerate() {
        assert_eq!(
            texts[cell],
            (0..count)
                .map(|index| format!("{prefix}{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn a_short_bottom_aligned_cell_stays_on_the_last_row_fragment() {
    assert_short_aligned_cells(&["top", "top", "bottom"]);
}

#[test]
fn short_centred_and_bottom_cells_keep_their_whole_row_alignment() {
    assert_short_aligned_cells(&["top", "top", "center", "bottom"]);
}

fn assert_short_aligned_cells(alignments: &[&str]) {
    let width = 8640 / alignments.len();
    let counts = [225, 216, 1, 1];
    let cells: String = alignments
        .iter()
        .enumerate()
        .map(|(cell, alignment)| {
            let spacing = if cell == 1 { 250 } else { 240 };
            let paragraphs: String = (0..counts[cell])
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7200_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("cell{cell}line{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{alignment}"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let grid = format!(r#"<w:gridCol w:w="{width}"/>"#).repeat(alignments.len());
    let bytes = fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ));
    let (engine, request) = fixture::laid_out(&bytes, 9404);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    let display: Value = serde_json::from_str(&display).unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 5);
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let row_height = output["measured"][0]["measure"]["rows"][0]["height"]
        .as_f64()
        .unwrap();
    let mut consumed = 0.0;
    let mut texts = vec![String::new(); alignments.len()];
    for (page_index, (page, painted)) in pages
        .iter()
        .zip(display["pages"].as_array().unwrap())
        .enumerate()
    {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        assert_eq!(fragments.len(), 1);
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            if cell > 1 {
                let expected_page = if alignments[cell] == "center" { 2 } else { 4 };
                assert_eq!(page_index, expected_page, "{text}");
            }
            let line = &output["measured"][0]["measure"]["rows"][0]["cells"][cell]["blocks"][0]["lines"]
                [0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_height = line["lineHeight"].as_f64().unwrap();
            let line_bottom = line_top + line_height;
            if cell > 1 {
                let slack = row_height - line_height;
                let offset = if alignments[cell] == "center" {
                    slack / 2.0
                } else {
                    slack
                };
                let expected_top = top + offset - consumed;
                assert!((line_top - expected_top).abs() < 0.02, "{text}: {line_top}");
            }
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
        consumed += fragment["height"].as_f64().unwrap();
    }
    for (cell, text) in texts.iter().enumerate() {
        assert_eq!(
            *text,
            (0..counts[cell])
                .map(|index| format!("cell{cell}line{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn an_exact_height_row_with_staggered_lines_stays_whole_and_clips() {
    let (output, display) = staggered_row_layout(true);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments: Vec<_> = pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(fragments.len(), 1);
    let fragment = fragments[0];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["height"], 480.0);
    assert_ne!(fragment["carriedToNext"], true);
    assert!(fragment["cellClips"].is_null());
    let top = fragment["y"].as_f64().unwrap();
    let bottom = top + fragment["height"].as_f64().unwrap();
    let mut text = String::new();
    for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
        let Some(run) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
            continue;
        };
        let clip = &primitive["clipGroup"]["clip"];
        assert!((clip["y"].as_f64().unwrap() - top).abs() < 0.02);
        assert!(clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap() <= bottom + 0.02);
        text.push_str(run);
    }
    assert!(text.contains("left000"));
    assert!(text.contains("right000"));
    assert!(!text.contains("left089"));
    assert!(!text.contains("right084"));
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
