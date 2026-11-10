use std::collections::BTreeMap;

use docx_edit::{
    EditApplication, EditCtx, EditFailureCode, EditRefusal, EditRequest, EditingDoc,
    InlineFormatDelta, Patch, SegmentContent, StoryInfoKind, StoryRange, StorySegment, UndoSession,
    seed_from_docx,
};
use serde_json::{Value, json};
use yrs::Any;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

const STYLES: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>"#;

fn package(body: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#;
    let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let document = format!(r#"<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    let styles = format!(r#"<w:styles {NS}>{STYLES}</w:styles>"#);
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            content_types.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), root_rels.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            document_rels.as_bytes().to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("word/styles.xml".to_owned(), styles.into_bytes()),
    ])
    .unwrap()
}

/// `00000001` plain, `00000002` directly bold, `00000003` a bold Heading1, `00000004` a run in the
/// bold character style `Strong`, `00000005` text around a line break, `00000007` a table cell
/// paragraph holding field characters inside a hyperlink, `00000009` a centred Heading1 run with direct bold-off, colour,
/// size and font, and `0000000A` an italic hyperlink.
fn open() -> EditingDoc {
    let body = concat!(
        r#"<w:p w14:paraId="00000001"><w:r><w:t>Plain text</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000002"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000003"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000004"><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>Strong run</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000005"><w:r><w:t>Line one</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>line two</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>"#,
        r#"<w:p w14:paraId="00000007"><w:r><w:t>Cell </w:t></w:r><w:hyperlink w:anchor="Target"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>"#,
        r#"</w:tc></w:tr></w:tbl>"#,
        r#"<w:p w14:paraId="00000008"><w:r><w:t>Last</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000009"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:color w:val="C00000"/><w:sz w:val="20"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Styled off</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="0000000A"><w:hyperlink w:anchor="Target"><w:r><w:rPr><w:i/></w:rPr><w:t>Linked</w:t></w:r></w:hyperlink></w:p>"#,
    );
    let doc = EditingDoc::new(7101);
    seed_from_docx(&doc, &package(body)).unwrap();
    doc
}

fn paragraph(id: &str) -> Value {
    json!({"kind": "paragraph", "story": "body", "paraId": id})
}

fn found(text: &str, id: &str) -> Value {
    json!({"kind": "search", "text": text, "view": "accepted",
        "within": {"kind": "paragraph", "story": "body", "paraId": id}})
}

fn cell_story(doc: &EditingDoc) -> String {
    doc.list_stories()
        .stories
        .into_iter()
        .find(|story| story.kind == StoryInfoKind::TableCell)
        .expect("a table cell story")
        .story
}

fn request(doc: &EditingDoc, history: &str, steps: Value) -> EditRequest {
    let json = json!({"expectVersion": doc.version().as_str(), "history": history, "steps": steps});
    EditRequest::from_json(&json.to_string())
        .unwrap()
        .unwrap_or_else(|failure| panic!("request refused while parsing: {failure:?}"))
}

fn apply(doc: &EditingDoc, undo: &UndoSession, steps: Value) -> EditApplication {
    doc.apply_edits(&request(doc, "separate", steps), undo)
        .unwrap()
        .unwrap_or_else(|refusal| panic!("batch refused: {refusal:?}"))
}

fn refuse(doc: &EditingDoc, steps: Value) -> EditRefusal {
    match doc
        .apply_edits(&request(doc, "separate", steps), &UndoSession::new())
        .unwrap()
    {
        Ok(applied) => panic!("batch applied: {applied:?}"),
        Err(refusal) => refusal,
    }
}

fn segments(doc: &EditingDoc, story: &str) -> Vec<StorySegment> {
    doc.story_segments(story).unwrap()
}

/// Attributes of the body text segment containing `marker`.
fn marks(doc: &EditingDoc, marker: &str) -> BTreeMap<String, Any> {
    marks_in(doc, "body", marker)
}

fn marks_in(doc: &EditingDoc, story: &str, marker: &str) -> BTreeMap<String, Any> {
    segments(doc, story)
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::Text(text) if text.contains(marker) => Some(segment.attributes),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text segment holds {marker:?}"))
}

fn active(map: &BTreeMap<String, Any>, key: &str) -> bool {
    map.get(key).is_some_and(|value| *value != Any::Null)
}

fn format(target: Value, extra: Value) -> Value {
    let mut step = json!({"op": "formatText", "target": target});
    for (key, value) in extra.as_object().unwrap() {
        step[key] = value.clone();
    }
    step
}

#[test]
fn clear_returns_text_to_what_its_styles_give() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([
            format(paragraph("00000002"), json!({"clear": ["bold"]})),
            format(paragraph("00000003"), json!({"clear": ["bold"]})),
            format(paragraph("00000004"), json!({"clear": ["bold"]})),
        ]),
    );
    assert!(!active(&marks(&doc, "Bold words"), "bold"));
    assert!(
        active(&marks(&doc, "Title"), "bold"),
        "a staged clear must resolve the paragraph style from the source"
    );
    let strong = marks(&doc, "Strong run");
    assert!(active(&strong, "bold") && active(&strong, "runStyle"));
}

#[test]
fn clear_all_keeps_the_paragraph_style_and_drops_the_character_style() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([
            format(paragraph("00000003"), json!({"clear": "all"})),
            format(paragraph("00000004"), json!({"clear": "all"})),
        ]),
    );
    let title = marks(&doc, "Title");
    assert!(active(&title, "bold") && active(&title, "fontSize"));
    let strong = marks(&doc, "Strong run");
    assert!(!active(&strong, "bold") && !active(&strong, "runStyle"));
}

#[test]
fn set_writes_the_attributes_the_formatting_commands_write() {
    let ctx = EditCtx::local("", "");
    let cases = [
        (
            json!({"bold": true, "italic": true}),
            InlineFormatDelta {
                bold: Patch::Set(true),
                italic: Patch::Set(true),
                ..Default::default()
            },
        ),
        (
            json!({"color": "C00000", "fontSizePt": 14}),
            InlineFormatDelta {
                color: Patch::Set(docx_edit::ColorPatch::Rgb("C00000".to_owned())),
                font_size: Patch::Set(14.0),
                ..Default::default()
            },
        ),
        (
            json!({"underline": "double", "highlight": "yellow"}),
            InlineFormatDelta {
                underline: Patch::Set(docx_edit::UnderlinePatch {
                    style: Some("double".to_owned()),
                    color: None,
                }),
                highlight: Patch::Set("yellow".to_owned()),
                ..Default::default()
            },
        ),
    ];
    for (set, delta) in cases {
        let step = open();
        apply(
            &step,
            &UndoSession::new(),
            json!([format(found("text", "00000001"), json!({"set": set}))]),
        );
        let command = open();
        command
            .format_range(&ctx, StoryRange::new("body", 6, 10), &delta)
            .unwrap();
        assert_eq!(segments(&step, "body"), segments(&command, "body"), "{set}");
    }
}

#[test]
fn false_turns_bold_off_in_every_context() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([
            format(paragraph("00000002"), json!({"set": {"bold": false}})),
            format(paragraph("00000003"), json!({"set": {"bold": false}})),
            format(paragraph("00000004"), json!({"set": {"bold": false}})),
        ]),
    );
    for marker in ["Bold words", "Title", "Strong run"] {
        assert!(!active(&marks(&doc, marker), "bold"), "{marker}");
    }
}

#[test]
fn targets_resolve_and_the_receipt_reports_the_range_with_text_unchanged() {
    let doc = open();
    let before: String = doc
        .read_paragraphs(&docx_edit::ReadParagraphsRequest {
            story: Some("body".to_owned()),
            para_ids: None,
            view: docx_edit::EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect();
    let range = json!({"kind": "range", "story": "body", "view": "accepted",
        "start": {"paraId": "00000008", "offset": 0}, "end": {"paraId": "00000008", "offset": 2}});
    let applied = apply(
        &doc,
        &UndoSession::new(),
        json!([
            format(found("Plain", "00000001"), json!({"set": {"italic": true}})),
            format(paragraph("00000005"), json!({"set": {"italic": true}})),
            format(range, json!({"set": {"italic": true}})),
        ]),
    );
    let receipts = serde_json::to_value(&applied.receipts).unwrap();
    assert_eq!(receipts[0]["range"]["start"]["offset"], 0);
    assert_eq!(receipts[0]["range"]["end"]["offset"], 5);
    assert_eq!(receipts[2]["range"]["end"]["offset"], 2);
    assert!(active(&marks(&doc, "Plain"), "italic"));
    assert!(!active(&marks(&doc, " text"), "italic"));
    assert!(
        active(&marks(&doc, "Line one"), "italic") && active(&marks(&doc, "line two"), "italic")
    );
    assert!(active(&marks(&doc, "La"), "italic") && !active(&marks(&doc, "st"), "italic"));
    let after: String = doc
        .read_paragraphs(&docx_edit::ReadParagraphsRequest {
            story: Some("body".to_owned()),
            para_ids: None,
            view: docx_edit::EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect();
    assert_eq!(after, before);
}

#[test]
fn kept_unchanged_paragraphs_refuse_by_default_and_are_listed_when_skipped() {
    let doc = open();
    let cell = cell_story(&doc);
    let cell_target = json!({"kind": "search", "text": "Cell", "view": "accepted",
        "within": {"kind": "story", "story": cell}});
    let before = doc.encode_state_as_update_v1();
    let refusal = refuse(
        &doc,
        json!([
            format(paragraph("00000001"), json!({"set": {"bold": true}})),
            format(cell_target.clone(), json!({"set": {"bold": true}})),
        ]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::Unsupported);
    assert_eq!(refusal.failure.step_index, Some(1));
    assert_eq!(doc.encode_state_as_update_v1(), before);

    let undo = UndoSession::new();
    let applied = apply(
        &doc,
        &undo,
        json!([
            format(
                paragraph("00000001"),
                json!({"set": {"bold": true}, "onUnsupported": "skip"})
            ),
            format(
                cell_target,
                json!({"set": {"bold": true}, "onUnsupported": "skip"})
            ),
        ]),
    );
    let receipts = serde_json::to_value(&applied.receipts).unwrap();
    assert_eq!(receipts[1]["changed"], false);
    assert_eq!(receipts[1]["skipped"][0]["reason"], "field");
    assert!(receipts[1]["skipped"][0]["paraId"].is_string());
    assert!(receipts[0].get("skipped").is_none());
    assert!(active(&marks(&doc, "Plain text"), "bold"));
    assert!(!active(&marks_in(&doc, &cell, "Cell"), "bold"));
    assert!(undo.undo());
    assert!(!undo.can_undo());
    assert!(!active(&marks(&doc, "Plain text"), "bold"));
}

#[test]
fn a_suggested_format_is_one_tracked_formatting_change() {
    let doc = open();
    let applied = apply(
        &doc,
        &UndoSession::new(),
        json!([format(
            paragraph("00000001"),
            json!({
                "set": {"bold": true},
                "suggest": {"author": "Ann", "date": "2026-11-10T12:00:00Z"}
            })
        )]),
    );
    assert_eq!(applied.receipts[0].revision_ids.len(), 1);
    let plain = marks(&doc, "Plain text");
    assert!(active(&plain, "bold") && active(&plain, "rPrChange"));
}

#[test]
fn formatting_and_inserting_are_one_undo_step() {
    let doc = open();
    let undo = UndoSession::new();
    let before = segments(&doc, "body");
    apply(
        &doc,
        &undo,
        json!([
            format(paragraph("00000001"), json!({"set": {"italic": true}})),
            {"op": "insertText", "target": paragraph("00000008"), "at": "end", "text": "!"}
        ]),
    );
    assert!(undo.undo());
    assert!(!undo.can_undo());
    assert_eq!(segments(&doc, "body"), before);
}

#[test]
fn formatting_that_overlaps_another_step_refuses() {
    let doc = open();
    let refusal = refuse(
        &doc,
        json!([
            format(found("Plain text", "00000001"), json!({"set": {"italic": true}})),
            {"op": "insertText", "target": found("text", "00000001"), "at": "start", "text": "x"}
        ]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::OverlappingSteps);
}

#[test]
fn invalid_format_steps_change_nothing() {
    let collapsed = json!({"kind": "range", "story": "body", "view": "accepted",
        "start": {"paraId": "00000001", "offset": 2}, "end": {"paraId": "00000001", "offset": 2}});
    let cases = [
        format(paragraph("00000001"), json!({})),
        format(paragraph("00000001"), json!({"set": {}})),
        format(paragraph("00000001"), json!({"clear": []})),
        format(
            paragraph("00000001"),
            json!({"set": {"bold": true}, "clear": ["bold"]}),
        ),
        format(paragraph("00000001"), json!({"set": {"link": "x"}})),
        format(paragraph("00000001"), json!({"set": {"color": "red"}})),
        format(paragraph("00000001"), json!({"clear": ["size"]})),
        format(collapsed, json!({"set": {"bold": true}})),
    ];
    for case in cases {
        let doc = open();
        let before = doc.encode_state_as_update_v1();
        let refusal = refuse(
            &doc,
            json!([
                {"op": "insertText", "target": paragraph("00000008"), "at": "end", "text": "!"},
                case.clone()
            ]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep, "{case}");
        assert_eq!(refusal.failure.step_index, Some(1), "{case}");
        assert_eq!(doc.encode_state_as_update_v1(), before, "{case}");
    }
    let doc = open();
    let json = json!({"expectVersion": doc.version().as_str(), "steps": [
        format(paragraph("00000001"), json!({"set": {"bold": true}, "onUnsupported": "ignore"}))
    ]});
    assert!(EditRequest::from_json(&json.to_string()).is_err());
}

#[test]
fn direct_and_staged_batches_format_the_same() {
    let steps = json!([format(
        paragraph("00000001"),
        json!({"set": {"bold": true, "color": "C00000"}})
    )]);
    let direct = open();
    direct.set_direct_batches(true);
    let staged = open();
    for doc in [&direct, &staged] {
        doc.apply_edits(&request(doc, "none", steps.clone()), &UndoSession::new())
            .unwrap()
            .unwrap();
    }
    assert_eq!(segments(&direct, "body"), segments(&staged, "body"));
}

#[test]
fn a_peer_converges_on_formatted_text() {
    let doc = open();
    let peer = EditingDoc::new(7102);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    let known = peer.encode_state_vector_v1();
    apply(
        &doc,
        &UndoSession::new(),
        json!([format(
            paragraph("00000002"),
            json!({"set": {"italic": true}, "clear": ["bold"]})
        )]),
    );
    peer.apply_update_v1(&doc.encode_diff_v1(&known).unwrap())
        .unwrap();
    assert_eq!(segments(&peer, "body"), segments(&doc, "body"));
}

#[test]
fn clearing_returns_toggle_and_value_marks_to_the_heading_style() {
    let doc = open();
    let styled = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([format(
            paragraph("00000009"),
            json!({"clear": ["bold", "color", "fontSizePt", "fontFamily"]})
        )]),
    );
    let cleared = marks(&doc, "Styled off");
    let title = marks(&styled, "Title");
    for key in [
        "bold",
        "fontSize",
        "textColor",
        "color",
        "fontFamily",
        "rFonts",
    ] {
        assert_eq!(cleared.get(key), title.get(key), "{key}");
    }
    assert!(active(&cleared, "bold"));
}

#[test]
fn undo_restores_the_direct_formatting_a_clear_removed() {
    let doc = open();
    let undo = UndoSession::new();
    let before = segments(&doc, "body");
    apply(
        &doc,
        &undo,
        json!([
            format(paragraph("00000009"), json!({"clear": "all"})),
            format(paragraph("00000002"), json!({"clear": ["bold"]})),
        ]),
    );
    assert_ne!(segments(&doc, "body"), before);
    assert!(undo.undo());
    assert!(!undo.can_undo());
    assert_eq!(segments(&doc, "body"), before);
}

#[test]
fn clear_all_leaves_links_and_paragraph_properties_alone() {
    let doc = open();
    let pilcrows = |doc: &EditingDoc| -> Vec<StorySegment> {
        segments(doc, "body")
            .into_iter()
            .filter(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
            .collect()
    };
    let paragraphs = pilcrows(&doc);
    let link = marks(&doc, "Linked").get("hyperlink").cloned();
    apply(
        &doc,
        &UndoSession::new(),
        json!([
            format(paragraph("00000009"), json!({"clear": "all"})),
            format(paragraph("0000000A"), json!({"clear": "all"})),
        ]),
    );
    assert_eq!(pilcrows(&doc), paragraphs);
    let linked = marks(&doc, "Linked");
    assert!(link.is_some());
    assert_eq!(linked.get("hyperlink").cloned(), link);
    assert!(!active(&linked, "italic"));
}

#[test]
fn a_suggested_clear_is_one_tracked_formatting_change() {
    let doc = open();
    let applied = apply(
        &doc,
        &UndoSession::new(),
        json!([format(
            paragraph("00000002"),
            json!({
                "clear": ["bold"],
                "suggest": {"author": "Ann", "date": "2026-11-10T12:00:00Z"}
            })
        )]),
    );
    assert_eq!(applied.receipts[0].revision_ids.len(), 1);
    let cleared = marks(&doc, "Bold words");
    assert!(!active(&cleared, "bold") && active(&cleared, "rPrChange"));
}

#[test]
fn peers_clearing_overlapping_ranges_converge() {
    let left = open();
    let right = EditingDoc::new(7103);
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    let range = |start: u32, end: u32| {
        json!({"kind": "range", "story": "body", "view": "accepted",
            "start": {"paraId": "00000002", "offset": start}, "end": {"paraId": "00000002", "offset": end}})
    };
    let (left_known, right_known) = (
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1(),
    );
    apply(
        &left,
        &UndoSession::new(),
        json!([format(range(0, 6), json!({"clear": ["bold"]}))]),
    );
    apply(
        &right,
        &UndoSession::new(),
        json!([format(
            range(3, 10),
            json!({"clear": "all", "set": {"italic": true}})
        )]),
    );
    let (to_right, to_left) = (
        left.encode_diff_v1(&right_known).unwrap(),
        right.encode_diff_v1(&left_known).unwrap(),
    );
    right.apply_update_v1(&to_right).unwrap();
    left.apply_update_v1(&to_left).unwrap();
    assert_eq!(segments(&left, "body"), segments(&right, "body"));
}
