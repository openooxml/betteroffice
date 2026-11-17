use std::collections::BTreeMap;

use docx_edit::{
    ColorPatch, EditApplication, EditCtx, EditFailureCode, EditRefusal, EditRequest, EditingDoc,
    FontFamilyPatch, InlineFormatDelta, Patch, SegmentContent, SimpleFormat, StoryRange,
    StorySegment, StrikePatch, UnderlinePatch, UndoSession, seed_from_docx,
};
use serde_json::{Value, json};
use yrs::Any;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

const STYLES: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Raised"><w:name w:val="Raised"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Marked"><w:name w:val="Marked"/><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>"#;

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

/// `00000001` plain, `00000002` directly bold, `00000003` a bold Heading1, `00000004` a run in
/// the bold character style `Strong`, `00000005` in the superscript style `Raised`, `00000006` in
/// the highlighting style `Marked`.
fn open() -> EditingDoc {
    let body = concat!(
        r#"<w:p w14:paraId="00000001"><w:r><w:t>Plain text</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000002"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000003"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000004"><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>Strong run</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000005"><w:pPr><w:pStyle w:val="Raised"/></w:pPr><w:r><w:t>Up</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000006"><w:pPr><w:pStyle w:val="Marked"/></w:pPr><w:r><w:t>Lit</w:t></w:r></w:p>"#,
    );
    let doc = EditingDoc::new(7001);
    seed_from_docx(&doc, &package(body)).unwrap();
    doc
}

fn paragraph(id: &str) -> Value {
    json!({"kind": "paragraph", "story": "body", "paraId": id})
}

fn anchor(id: &str) -> Value {
    json!({"story": "body", "paraId": id})
}

fn found(text: &str, id: &str) -> Value {
    json!({"kind": "search", "text": text, "view": "accepted",
        "within": {"kind": "paragraph", "story": "body", "paraId": id}})
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

fn segments(doc: &EditingDoc) -> Vec<StorySegment> {
    doc.story_segments("body").unwrap()
}

fn text_of(doc: &EditingDoc) -> String {
    segments(doc)
        .into_iter()
        .filter_map(|segment| match segment.content {
            SegmentContent::Text(text) => Some(text),
            _ => None,
        })
        .collect()
}

/// Attributes of the text segment containing `marker`.
fn marks(doc: &EditingDoc, marker: &str) -> BTreeMap<String, Any> {
    segments(doc)
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

/// The step with `text` replaced by `runs`.
fn with_runs(mut step: Value, runs: Value) -> Value {
    let object = step.as_object_mut().unwrap();
    object.remove("text");
    object.insert("runs".to_owned(), runs);
    step
}

#[test]
fn runs_without_marks_write_exactly_what_plain_text_writes() {
    let steps = [
        json!({"op": "insertText", "target": paragraph("00000001"), "at": "end", "text": " x\ty"}),
        json!({"op": "insertText", "target": paragraph("00000002"), "at": "end", "text": " x\ty"}),
        json!({"op": "insertText", "target": paragraph("00000002"), "at": "start", "text": "x "}),
        json!({"op": "insertText", "target": paragraph("00000003"), "at": "end", "text": " x"}),
        json!({"op": "replaceText", "target": found("words", "00000002"), "text": "terms"}),
        json!({"op": "replaceText", "target": found("Plain", "00000001"), "text": "Simple"}),
    ];
    for step in steps {
        let plain = open();
        let runs = open();
        let text = step["text"].clone();
        apply(&plain, &UndoSession::new(), json!([step.clone()]));
        apply(
            &runs,
            &UndoSession::new(),
            json!([with_runs(step.clone(), json!([{"text": text}]))]),
        );
        assert_eq!(segments(&runs), segments(&plain), "{step}");
    }
}

#[test]
fn several_formatted_runs_and_another_step_are_one_undo_step() {
    let doc = open();
    let undo = UndoSession::new();
    let before = segments(&doc);
    let applied = apply(
        &doc,
        &undo,
        json!([
            {"op": "insertText", "target": paragraph("00000001"), "at": "end", "runs": [
                {"text": " B1", "marks": {"bold": true}},
                {"text": " I2", "marks": {"italic": true}},
                {"text": " P3"}
            ]},
            {"op": "replaceText", "target": found("Title", "00000003"), "text": "Heading"}
        ]),
    );
    assert!(applied.applied);
    assert_eq!(
        text_of(&doc),
        "Plain text B1 I2 P3Bold wordsHeadingStrong runUpLit"
    );
    let bold = marks(&doc, "B1");
    assert!(active(&bold, "bold") && !active(&bold, "italic"));
    let italic = marks(&doc, "I2");
    assert!(active(&italic, "italic") && !active(&italic, "bold"));
    let plain = marks(&doc, "P3");
    assert!(!active(&plain, "bold") && !active(&plain, "italic"));
    assert!(undo.undo());
    assert!(!undo.can_undo());
    assert_eq!(segments(&doc), before);
}

#[test]
fn runs_overlay_marks_on_the_formatting_at_the_insertion_point() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([{"op": "insertText", "target": paragraph("00000002"), "at": "end", "runs": [
            {"text": " keep"},
            {"text": " off", "marks": {"bold": false}},
            {"text": " both", "marks": {"italic": true}}
        ]}]),
    );
    assert!(active(&marks(&doc, "keep"), "bold"));
    let off = marks(&doc, "off");
    assert!(!active(&off, "bold") && !active(&off, "italic"));
    let both = marks(&doc, "both");
    assert!(active(&both, "bold") && active(&both, "italic"));
}

#[test]
fn marks_write_the_attributes_the_formatting_commands_write() {
    let ctx = EditCtx::local("", "");
    let delta = |delta: InlineFormatDelta| Some(delta);
    let cases: Vec<(Value, Option<InlineFormatDelta>, Option<SimpleFormat>)> = vec![
        (
            json!({"bold": true}),
            delta(InlineFormatDelta {
                bold: Patch::Set(true),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"underline": true}),
            delta(InlineFormatDelta {
                underline: Patch::Set(UnderlinePatch::default()),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"underline": "double"}),
            delta(InlineFormatDelta {
                underline: Patch::Set(UnderlinePatch {
                    style: Some("double".to_owned()),
                    color: None,
                }),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"strike": "dblStrike"}),
            delta(InlineFormatDelta {
                strike: Patch::Set(StrikePatch { double: true }),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"strike": true}),
            delta(InlineFormatDelta {
                strike: Patch::Set(StrikePatch { double: false }),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"color": "C00000"}),
            delta(InlineFormatDelta {
                color: Patch::Set(ColorPatch::Rgb("C00000".to_owned())),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"color": {"themeColor": "accent1"}}),
            delta(InlineFormatDelta {
                color: Patch::Set(ColorPatch::Theme("accent1".to_owned())),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"highlight": "yellow"}),
            delta(InlineFormatDelta {
                highlight: Patch::Set("yellow".to_owned()),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"fontSizePt": 10.5}),
            delta(InlineFormatDelta {
                font_size: Patch::Set(10.5),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"fontFamily": "Arial"}),
            delta(InlineFormatDelta {
                font_family: Patch::Set(FontFamilyPatch {
                    ascii: "Arial".to_owned(),
                    h_ansi: Some("Arial".to_owned()),
                }),
                ..Default::default()
            }),
            None,
        ),
        (
            json!({"verticalAlign": "superscript"}),
            None,
            Some(SimpleFormat::Superscript),
        ),
        (
            json!({"verticalAlign": "subscript"}),
            None,
            Some(SimpleFormat::Subscript),
        ),
    ];
    for (mark, delta, toggle) in cases {
        let runs = open();
        apply(
            &runs,
            &UndoSession::new(),
            json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end",
                "runs": [{"text": "Z", "marks": mark}]}]),
        );
        let command = open();
        apply(
            &command,
            &UndoSession::new(),
            json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end", "text": "Z"}]),
        );
        let range = StoryRange::new("body", 10, 11);
        if let Some(delta) = delta {
            command.format_range(&ctx, range, &delta).unwrap();
        } else if let Some(toggle) = toggle {
            command.toggle_format(&ctx, range, toggle).unwrap();
        }
        assert_eq!(marks(&runs, "Z"), marks(&command, "Z"), "{mark}");
    }
}

#[test]
fn refused_runs_and_marks_leave_the_document_untouched() {
    let valid =
        json!({"op": "insertText", "target": paragraph("00000003"), "at": "end", "text": "!"});
    let runs = |runs: Value| json!({"op": "insertText", "target": paragraph("00000001"), "at": "end", "runs": runs});
    let marked = |marks: Value| runs(json!([{"text": "a", "marks": marks}]));
    let cases = [
        runs(json!([])),
        runs(json!([{"text": ""}])),
        runs(json!([{"text": "a\nb"}])),
        marked(json!({"link": {"href": "https://example.com"}})),
        marked(json!({"bold": "yes"})),
        marked(json!({"color": "red"})),
        marked(json!({"color": {"themeColor": "accent9"}})),
        marked(json!({"highlight": "orange"})),
        marked(json!({"underline": "zigzag"})),
        marked(json!({"strike": "double"})),
        marked(json!({"verticalAlign": "up"})),
        marked(json!({"verticalAlign": "baseline"})),
        marked(json!({"highlight": "none"})),
        json!({"op": "insertText", "target": paragraph("00000005"), "at": "end",
            "runs": [{"text": "a", "marks": {"verticalAlign": "baseline"}}]}),
        json!({"op": "insertText", "target": paragraph("00000006"), "at": "end",
            "runs": [{"text": "a", "marks": {"highlight": "none"}}]}),
        marked(json!({"fontSizePt": 0.5})),
        marked(json!({"fontSizePt": 10.25})),
        marked(json!({"fontFamily": ""})),
        json!({"op": "insertParagraphs", "target": anchor("00000001"), "at": "end",
            "paragraphs": [{"runs": [{"text": "a", "marks": {"size": 12}}]}]}),
    ];
    for case in cases {
        let doc = open();
        let before = doc.encode_state_as_update_v1();
        let version = doc.version();
        let refusal = refuse(&doc, json!([valid.clone(), case.clone()]));
        assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep, "{case}");
        assert_eq!(refusal.failure.step_index, Some(1), "{case}");
        assert_eq!(doc.encode_state_as_update_v1(), before, "{case}");
        assert_eq!(doc.version(), version, "{case}");
    }
}

#[test]
fn a_step_takes_exactly_one_of_text_and_runs() {
    let doc = open();
    for step in [
        json!({"op": "insertText", "target": paragraph("00000001"), "at": "end", "text": "a", "runs": [{"text": "a"}]}),
        json!({"op": "insertText", "target": paragraph("00000001"), "at": "end"}),
        json!({"op": "replaceText", "target": found("Plain", "00000001"), "runs": null}),
        json!({"op": "insertParagraphs", "target": anchor("00000001"), "at": "end",
            "paragraphs": [{"text": "a", "runs": [{"text": "a"}]}]}),
        json!({"op": "insertParagraphs", "target": anchor("00000001"), "at": "end",
            "paragraphs": [{"styleId": "Normal"}]}),
        json!({"op": "insertText", "target": paragraph("00000001"), "at": "end",
            "runs": [{"text": "a", "bold": true}]}),
    ] {
        let json = json!({"expectVersion": doc.version().as_str(), "steps": [step]});
        assert!(EditRequest::from_json(&json.to_string()).is_err(), "{json}");
    }
}

#[test]
fn adjacent_runs_with_equal_marks_read_as_one_run() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end", "runs": [
            {"text": " aa", "marks": {"bold": true}},
            {"text": "bb", "marks": {"bold": true}}
        ]}]),
    );
    assert!(active(&marks(&doc, " aabb"), "bold"));
}

#[test]
fn a_combining_sequence_split_across_runs_keeps_its_text() {
    let doc = open();
    apply(
        &doc,
        &UndoSession::new(),
        json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end", "runs": [
            {"text": " e", "marks": {"bold": true}},
            {"text": "\u{301}", "marks": {"italic": true}}
        ]}]),
    );
    assert!(text_of(&doc).starts_with("Plain text e\u{301}Bold"));
    assert!(active(&marks(&doc, "\u{301}"), "italic"));
}

#[test]
fn suggested_runs_share_one_revision() {
    let doc = open();
    let applied = apply(
        &doc,
        &UndoSession::new(),
        json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end",
            "suggest": {"author": "Ann", "date": "2026-11-10T12:00:00Z"},
            "runs": [{"text": " B1", "marks": {"bold": true}}, {"text": " P2"}]}]),
    );
    assert_eq!(applied.receipts[0].revision_ids.len(), 1);
    let bold = marks(&doc, "B1");
    let plain = marks(&doc, "P2");
    assert!(active(&bold, "bold") && !active(&plain, "bold"));
    assert!(active(&bold, "ins") && bold.get("ins") == plain.get("ins"));
}

#[test]
fn new_paragraph_runs_overlay_the_paragraph_style() {
    let plain = open();
    apply(
        &plain,
        &UndoSession::new(),
        json!([{"op": "insertParagraphs", "target": anchor("00000001"), "at": "end",
            "paragraphs": [{"text": "Head plain", "styleId": "Heading1"}]}]),
    );
    let runs = open();
    apply(
        &runs,
        &UndoSession::new(),
        json!([{"op": "insertParagraphs", "target": anchor("00000001"), "at": "end",
        "paragraphs": [{"styleId": "Heading1", "runs": [
            {"text": "Head", "marks": {"italic": true}},
            {"text": " plain"}
        ]}]}]),
    );
    let base = marks(&plain, "Head plain");
    assert_eq!(marks(&runs, " plain"), base);
    let mut italic = base.clone();
    italic.insert("italic".to_owned(), Any::Bool(true));
    assert_eq!(marks(&runs, "Head"), italic);
    assert_eq!(text_of(&runs), text_of(&plain));
}

#[test]
fn direct_and_staged_batches_write_the_same_runs() {
    let steps = json!([{"op": "insertText", "target": paragraph("00000001"), "at": "end", "runs": [
        {"text": " B1", "marks": {"bold": true}},
        {"text": " P2"}
    ]}]);
    let direct = open();
    direct.set_direct_batches(true);
    let staged = open();
    for doc in [&direct, &staged] {
        doc.apply_edits(&request(doc, "none", steps.clone()), &UndoSession::new())
            .unwrap()
            .unwrap();
    }
    assert_eq!(segments(&direct), segments(&staged));
}

#[test]
fn a_peer_converges_on_formatted_runs() {
    let doc = open();
    let peer = EditingDoc::new(7002);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    let known = peer.encode_state_vector_v1();
    apply(
        &doc,
        &UndoSession::new(),
        json!([{"op": "insertText", "target": paragraph("00000002"), "at": "end", "runs": [
            {"text": " I1", "marks": {"italic": true, "bold": false}},
            {"text": " C2", "marks": {"color": "C00000"}}
        ]}]),
    );
    peer.apply_update_v1(&doc.encode_diff_v1(&known).unwrap())
        .unwrap();
    assert_eq!(segments(&peer), segments(&doc));
    assert_eq!(
        peer.encode_state_as_update_v1(),
        doc.encode_state_as_update_v1()
    );
}

#[test]
fn false_turns_a_mark_off_whether_it_is_direct_or_from_a_style() {
    for id in ["00000002", "00000003", "00000004"] {
        let doc = open();
        apply(
            &doc,
            &UndoSession::new(),
            json!([{"op": "insertText", "target": paragraph(id), "at": "end", "runs": [
                {"text": " kept"},
                {"text": " off", "marks": {"bold": false}}
            ]}]),
        );
        assert!(active(&marks(&doc, " kept"), "bold"), "{id}");
        assert!(!active(&marks(&doc, " off"), "bold"), "{id}");
    }
}

#[test]
fn edit_operations_deserialized_on_their_own_take_exactly_one_of_text_and_runs() {
    let target = paragraph("00000001");
    for operation in [
        json!({"op": "replaceText", "target": target}),
        json!({"op": "insertText", "target": target, "at": "end"}),
        json!({"op": "insertText", "target": target, "at": "end", "text": "a", "runs": [{"text": "a"}]}),
        json!({"op": "insertParagraphs", "target": {"story": "body", "paraId": "00000001"}, "at": "end",
            "paragraphs": [{"styleId": "Normal"}]}),
    ] {
        assert!(
            serde_json::from_value::<docx_edit::EditOperation>(operation.clone()).is_err(),
            "{operation}"
        );
    }
    assert!(
        serde_json::from_value::<docx_edit::ParagraphInput>(json!({"styleId": "Normal"})).is_err()
    );
}
