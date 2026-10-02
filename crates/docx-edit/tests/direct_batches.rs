use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex};

use docx_edit::{
    EditApplication, EditCtx, EditFailureCode, EditHistory, EditOperation, EditRefusal,
    EditRequest, EditSource, EditStep, EditSuggestion, EditTextView, EditingDoc, FormatPolicy,
    ParagraphInput, ParagraphTarget, SearchScope, SegmentContent, TargetEdge, TextPosition,
    TextRange, TextTarget, UndoSession, seed_from_docx_with_generation,
};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use serde_json::{Value, json};
use yrs::updates::decoder::Decode;
use yrs::{Any, DeepObservable, Map, Origin, ReadTxn, Transact};

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
    r#"xmlns:bofx="urn:fidelity""#
);

const STYLES: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:caps/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListChild"><w:name w:val="List Child"/><w:basedOn w:val="ListNumber"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListGrandchild"><w:name w:val="List Grandchild"/><w:basedOn w:val="ListChild"/></w:style><w:style w:type="paragraph" w:styleId="QuoteChild"><w:name w:val="Quote Child"/><w:basedOn w:val="Quote"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>"#;

const DATE: &str = "2026-09-24T12:00:00Z";

fn docx(body: &str) -> Vec<u8> {
    with_comment(body, "<w:p><w:r><w:t>Remark</w:t></w:r></w:p>")
}

fn with_comment(body: &str, comment: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>"#;
    let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let rel = |id: &str, kind: &str, target: &str| {
        format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        )
    };
    let document_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}{}{}{}{}</Relationships>"#,
        rel("rIdStyles", "styles", "styles.xml"),
        rel("rIdHeader", "header", "header1.xml"),
        rel("rIdNotes", "footnotes", "footnotes.xml"),
        rel("rIdComments", "comments", "comments.xml"),
        rel("rIdImage", "image", "media/one.png"),
    );
    let document = format!(
        r#"<w:document {NS}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>"#
    );
    let header = format!(
        r#"<w:hdr {NS}><w:p w14:paraId="0000E001"><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>"#
    );
    let notes = format!(
        r#"<w:footnotes {NS}><w:footnote w:id="1"><w:p><w:r><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    );
    let comments = format!(
        r#"<w:comments {NS}><w:comment w:id="9" w:author="Ann" w:date="{DATE}">{comment}</w:comment></w:comments>"#
    );
    let styles = format!(r#"<w:styles {NS}>{STYLES}</w:styles>"#);
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            content_types.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), root_rels.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            document_rels.into_bytes(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("word/styles.xml".to_owned(), styles.into_bytes()),
        ("word/header1.xml".to_owned(), header.into_bytes()),
        ("word/footnotes.xml".to_owned(), notes.into_bytes()),
        ("word/comments.xml".to_owned(), comments.into_bytes()),
        ("word/media/one.png".to_owned(), vec![137, 80, 78, 71]),
    ])
    .unwrap()
}

fn p(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#)
}

fn r(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn state_vector(doc: &EditingDoc) -> yrs::StateVector {
    yrs::StateVector::decode_v1(&doc.encode_state_vector_v1()).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(7001);
    seed_from_docx_with_generation(&doc, bytes, "direct-batches").unwrap();
    doc
}

fn fixture() -> Vec<u8> {
    docx(&format!(
        "{}{}{}{}{}",
        p("00000001", &r("Alpha beta gamma")),
        p("00000002", &r("Delta epsilon")),
        p("00000003", ""),
        p("00000004", &r("User")),
        p("00000005", &r("Probe")),
    ))
}

fn target(story: &str, id: &str, start: u32, end: u32) -> TextTarget {
    TextTarget::Range(TextRange {
        story: story.to_owned(),
        start: TextPosition {
            para_id: id.to_owned(),
            offset: start,
        },
        end: TextPosition {
            para_id: id.to_owned(),
            offset: end,
        },
        view: EditTextView::Accepted,
    })
}

fn range(id: &str, start: u32, end: u32) -> TextTarget {
    target("body", id, start, end)
}

fn search(id: &str, text: &str) -> TextTarget {
    TextTarget::Search {
        text: text.to_owned(),
        within: SearchScope::Paragraph(ParagraphTarget {
            story: "body".to_owned(),
            para_id: id.to_owned(),
        }),
        view: EditTextView::Accepted,
    }
}

fn insert(target: TextTarget, text: &str) -> EditStep {
    EditStep::new(EditOperation::InsertText {
        target,
        at: TargetEdge::Start,
        text: text.to_owned(),
    })
}

fn replace(target: TextTarget, text: &str) -> EditStep {
    EditStep::new(EditOperation::ReplaceText {
        target,
        text: text.to_owned(),
    })
}

fn delete(target: TextTarget) -> EditStep {
    EditStep::new(EditOperation::DeleteText { target })
}

fn suggested(mut step: EditStep) -> EditStep {
    step.suggest = Some(EditSuggestion {
        author: "Ann".to_owned(),
        date: DATE.to_owned(),
    });
    step
}

fn request(doc: &EditingDoc, steps: Vec<EditStep>, history: EditHistory) -> EditRequest {
    EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history,
        steps,
    }
}

fn apply(doc: &EditingDoc, undo: &UndoSession, steps: Vec<EditStep>) -> EditApplication {
    apply_request(doc, &request(doc, steps, EditHistory::None), undo).unwrap()
}

fn apply_request(
    doc: &EditingDoc,
    request: &EditRequest,
    undo: &UndoSession,
) -> Result<EditApplication, EditRefusal> {
    let base_version = doc.version();
    let epoch = doc.committed_epoch();
    let outcome = doc.apply_edits(request, undo).unwrap();
    let applied = match &outcome {
        Ok(result) => {
            assert_eq!(result.base_version, base_version);
            assert_eq!(result.version, doc.version());
            assert_eq!(result.version != base_version, result.applied);
            result.applied
        }
        Err(refusal) => {
            assert_eq!(refusal.version, base_version);
            assert_eq!(refusal.version, doc.version());
            false
        }
    };
    assert_eq!(doc.committed_epoch() - epoch, u64::from(applied));
    outcome
}

fn normalized<T: serde::Serialize>(outcome: &T) -> Value {
    let mut value = serde_json::to_value(outcome).unwrap();
    for key in ["baseVersion", "version"] {
        if value.get(key).is_some() {
            value[key] = json!("version");
        }
    }
    value
}

fn story_ids(doc: &EditingDoc) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    let mut ids: Vec<_> = txn
        .get_map("stories")
        .unwrap()
        .keys(&txn)
        .map(str::to_owned)
        .collect();
    ids.sort();
    ids
}

fn assert_stories(left: &EditingDoc, right: &EditingDoc) {
    let ids = story_ids(left);
    assert_eq!(ids, story_ids(right));
    for story in ids {
        assert_eq!(
            left.story_segments(&story).unwrap(),
            right.story_segments(&story).unwrap(),
            "{story}"
        );
        assert_eq!(
            left.paragraphs(&story)
                .unwrap()
                .into_iter()
                .map(|p| p.para_id)
                .collect::<Vec<_>>(),
            right
                .paragraphs(&story)
                .unwrap()
                .into_iter()
                .map(|p| p.para_id)
                .collect::<Vec<_>>(),
        );
    }
    assert_eq!(
        left.paragraph_identities().paragraphs,
        right.paragraph_identities().paragraphs
    );
}

#[derive(Clone)]
struct Observed {
    update: Vec<u8>,
    stories: BTreeSet<String>,
    origin: Option<Origin>,
}

struct Updates {
    events: Arc<Mutex<Vec<Observed>>>,
    _deep: yrs::Subscription,
    _updates: yrs::Subscription,
}

impl Updates {
    fn new(doc: &EditingDoc) -> Self {
        let touched = Arc::new(Mutex::new(BTreeSet::new()));
        let seen = Arc::clone(&touched);
        let stories = doc.yrs_doc().transact().get_map("stories").unwrap();
        let deep = stories.observe_deep(move |txn, events| {
            for event in events.iter() {
                match event.path().front() {
                    Some(yrs::types::PathSegment::Key(story)) => {
                        seen.lock().unwrap().insert(story.to_string());
                    }
                    None => {
                        if let yrs::types::Event::Map(event) = event {
                            seen.lock()
                                .unwrap()
                                .extend(event.keys(txn).keys().map(|id| id.to_string()));
                        }
                    }
                    _ => {}
                }
            }
        });
        let events = Arc::new(Mutex::new(Vec::new()));
        let captured = Arc::clone(&events);
        let updates = doc
            .yrs_doc()
            .observe_update_v1(move |txn, event| {
                captured.lock().unwrap().push(Observed {
                    update: event.update.clone(),
                    stories: std::mem::take(&mut *touched.lock().unwrap()),
                    origin: txn.origin().cloned(),
                });
            })
            .unwrap();
        Self {
            events,
            _deep: deep,
            _updates: updates,
        }
    }

    fn signature(&self) -> Vec<BTreeSet<String>> {
        self.events
            .lock()
            .unwrap()
            .iter()
            .map(|event| event.stories.clone())
            .collect()
    }
}

fn any_json(any: &Any) -> Value {
    let mut json = String::new();
    any.to_json(&mut json);
    serde_json::from_str(&json).unwrap()
}

fn export_story(
    doc: &EditingDoc,
    story: &str,
    revisions: &BTreeMap<(String, String), u32>,
) -> Value {
    let mut paragraphs = Vec::new();
    let mut content = Vec::new();
    let identities = doc.paragraph_identities().paragraphs;
    for segment in doc.story_segments(story).unwrap() {
        match segment.content {
            SegmentContent::Text(text) => {
                let formatting: serde_json::Map<_, _> = ["bold", "italic"]
                    .into_iter()
                    .filter_map(|key| {
                        segment
                            .attributes
                            .get(key)
                            .map(|value| (key.to_owned(), any_json(value)))
                    })
                    .filter(|(_, value)| !value.is_null())
                    .collect();
                let run = json!({"type": "run", "formatting": formatting,
                    "content": [{"type": "text", "text": text, "preserveSpace": true}]});
                let tracked = ["ins", "del"].into_iter().find_map(|key| {
                    let value = segment.attributes.get(key)?;
                    if matches!(value, Any::Null | Any::Undefined) {
                        return None;
                    }
                    Some((key, any_json(value)))
                });
                content.push(if let Some((key, revision)) = tracked {
                    let id = revision["id"].as_str().unwrap();
                    json!({"type": if key == "ins" { "insertion" } else { "deletion" },
                        "info": {"id": revisions[&(id.to_owned(), key.to_owned())],
                            "author": revision["author"], "date": revision["date"]},
                        "content": [run]})
                } else {
                    run
                });
            }
            SegmentContent::Pilcrow(properties) => {
                let id = identities
                    .iter()
                    .find_map(|identity| match &identity.paragraph {
                        docx_edit::ParagraphRef::Session { story: id, para_id }
                            if id == story && para_id == &properties.para_id =>
                        {
                            identity.ooxml_para_id.clone()
                        }
                        _ => None,
                    });
                let formatting = properties
                    .values
                    .get("formatting")
                    .map(any_json)
                    .unwrap_or_else(|| json!({}));
                paragraphs.push(json!({"type": "paragraph", "paraId": id,
                    "formatting": formatting, "content": std::mem::take(&mut content)}));
            }
            SegmentContent::OtherEmbed { kind, .. } => panic!("unexpected export embed {kind}"),
        }
    }
    assert!(content.is_empty());
    json!(paragraphs)
}

fn export(doc: &EditingDoc, bytes: &[u8]) -> Vec<u8> {
    let mut revisions = BTreeSet::new();
    for story in story_ids(doc) {
        for segment in doc.story_segments(&story).unwrap() {
            for key in ["ins", "del"] {
                if let Some(Any::Map(revision)) = segment.attributes.get(key)
                    && let Some(Any::String(id)) = revision.get("id")
                {
                    revisions.insert((id.to_string(), key.to_owned()));
                }
            }
        }
    }
    let revisions: BTreeMap<_, _> = revisions
        .into_iter()
        .enumerate()
        .map(|(i, id)| (id, i as u32))
        .collect();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut body = serde_json::to_value(&package.document).unwrap();
    body.as_object_mut().unwrap().remove("sections");
    body["content"] = export_story(doc, "body", &revisions);
    let mut headers = serde_json::to_value(package.header_entries.unwrap_or_default()).unwrap();
    headers[0][1]["content"] = export_story(doc, "hf:rIdHeader", &revisions);
    let mut notes = serde_json::to_value(package.footnotes.unwrap_or_default()).unwrap();
    if story_ids(doc).iter().any(|story| story == "fn:1") {
        notes[0]["content"] = export_story(doc, "fn:1", &revisions);
    }
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": body, "headerEntries": headers, "footnotes": notes,
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
    }))
    .unwrap();
    write_docx_s13(request, bytes).unwrap()
}

fn assert_export(left: &EditingDoc, right: &EditingDoc, bytes: &[u8]) {
    assert_eq!(export(left, bytes), export(right, bytes));
}

fn exercise(bytes: &[u8], prepare: impl Fn(&EditingDoc), steps: Vec<EditStep>) -> EditApplication {
    exercise_path(bytes, prepare, steps, true)
}

fn exercise_path(
    bytes: &[u8],
    prepare: impl Fn(&EditingDoc),
    steps: Vec<EditStep>,
    direct_path: bool,
) -> EditApplication {
    let direct = open(bytes);
    let replica = open(bytes);
    for doc in [&direct, &replica] {
        prepare(doc);
    }
    let direct_undo = UndoSession::with_clock(Arc::new(|| 1_000));
    let replica_undo = UndoSession::with_clock(Arc::new(|| 1_000));
    let direct_depth = direct.undo_manager();
    let replica_depth = replica.undo_manager();
    for (doc, undo) in [(&direct, &direct_undo), (&replica, &replica_undo)] {
        undo.track(doc);
        doc.insert_text(
            &EditCtx::local("User", DATE),
            doc.paragraph_mark_position("00000004").unwrap(),
            "!",
            FormatPolicy::Inherit,
        )
        .unwrap();
    }
    let base = direct.encode_state_as_update_v1();
    assert_eq!(base, replica.encode_state_as_update_v1());
    let peer = EditingDoc::new(7002);
    peer.apply_update_v1(&base).unwrap();
    peer.retain_source_docx(bytes.to_vec());
    let known = peer.encode_state_vector_v1();
    let updates_direct = Updates::new(&direct);
    let updates_replica = Updates::new(&replica);
    direct.set_direct_batches(true);
    direct.begin_shared_reads();
    replica.begin_shared_reads();
    let result_direct = apply(&direct, &direct_undo, steps.clone());
    let result_replica = apply(&replica, &replica_undo, steps);
    direct.end_shared_reads();
    replica.end_shared_reads();
    assert!(result_direct.applied);
    assert_eq!(direct.direct_batches_applied(), u64::from(direct_path));
    assert_eq!(replica.direct_batches_applied(), 0);
    assert_eq!(normalized(&result_direct), normalized(&result_replica));
    assert_eq!(state_vector(&direct), state_vector(&replica));
    assert_eq!(updates_direct.signature(), updates_replica.signature());
    assert_eq!(updates_direct.events.lock().unwrap().len(), 1);
    assert_eq!(
        updates_direct.events.lock().unwrap()[0].origin,
        Some(Origin::from("host"))
    );
    assert_eq!(
        updates_direct.events.lock().unwrap()[0].stories,
        result_direct
            .changed_stories
            .iter()
            .cloned()
            .collect::<BTreeSet<_>>()
    );
    assert_eq!(
        (direct_depth.undo_depth(), direct_depth.redo_depth()),
        (1, 0)
    );
    assert_eq!(
        (direct_depth.undo_depth(), direct_depth.redo_depth()),
        (replica_depth.undo_depth(), replica_depth.redo_depth())
    );
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, bytes);
    assert!(direct_undo.undo());
    assert!(replica_undo.undo());
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, bytes);
    assert!(direct_undo.redo());
    assert!(replica_undo.redo());
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, bytes);
    for id in ["00000001", "00000002"] {
        peer.insert_text(
            &EditCtx::local("Peer", DATE),
            peer.paragraph_mark_position(id).unwrap(),
            " remote",
            FormatPolicy::Inherit,
        )
        .unwrap();
    }
    let remote = peer.encode_diff_v1(&known).unwrap();
    direct.apply_update_v1(&remote).unwrap();
    replica.apply_update_v1(&remote).unwrap();
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, bytes);
    direct.set_direct_batches(false);
    let followup = vec![suggested(insert(range("00000005", 0, 0), "next "))];
    let followup_direct = apply(&direct, &direct_undo, followup.clone());
    let followup_replica = apply(&replica, &replica_undo, followup);
    assert_eq!(followup_direct.receipts[0].revision_ids.len(), 1);
    assert_eq!(normalized(&followup_direct), normalized(&followup_replica));
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, bytes);
    assert_eq!(updates_direct.signature(), updates_replica.signature());
    result_direct
}

#[test]
fn plain_and_suggested_text_operations() {
    for step in [
        insert(range("00000001", 6, 6), "new "),
        replace(search("00000001", "beta"), "BETA"),
        delete(search("00000001", "beta")),
    ] {
        for suggest in [false, true] {
            let step = if suggest {
                suggested(step.clone())
            } else {
                step.clone()
            };
            exercise(&fixture(), |_| {}, vec![step]);
        }
    }
}

#[test]
fn descending_steps_in_one_paragraph() {
    let steps = vec![
        suggested(insert(range("00000001", 0, 0), "First ")),
        suggested(replace(search("00000001", "beta"), "BETA!")),
        suggested(delete(search("00000001", "gamma"))),
    ];
    exercise(&fixture(), |_| {}, steps.clone());
    let plain: Vec<_> = steps
        .into_iter()
        .map(|mut step| {
            step.suggest = None;
            step
        })
        .collect();
    exercise(&fixture(), |_| {}, plain[..2].to_vec());
    exercise_path(&fixture(), |_| {}, plain, false);
}

#[test]
fn steps_in_two_paragraphs_and_two_stories() {
    exercise(
        &fixture(),
        |_| {},
        vec![
            suggested(replace(search("00000001", "beta"), "BETA")),
            delete(search("00000002", "epsilon")),
            suggested(insert(target("hf:rIdHeader", "0000E001", 7, 7), "new ")),
        ],
    );
}

#[test]
fn existing_suggestions_keep_their_planning_refusals() {
    for step in [
        suggested(insert(range("00000001", 5, 5), "new")),
        suggested(delete(search("00000001", "owned"))),
    ] {
        let bytes = fixture();
        let direct = open(&bytes);
        let replica = open(&bytes);
        for doc in [&direct, &replica] {
            apply(
                doc,
                &UndoSession::new(),
                vec![suggested(insert(range("00000001", 0, 0), "owned"))],
            );
        }
        let base = direct.encode_state_as_update_v1();
        assert_eq!(base, replica.encode_state_as_update_v1());
        direct.set_direct_batches(true);
        let a = apply_request(
            &direct,
            &request(&direct, vec![step.clone()], EditHistory::None),
            &UndoSession::new(),
        )
        .unwrap_err();
        let b = apply_request(
            &replica,
            &request(&replica, vec![step], EditHistory::None),
            &UndoSession::new(),
        )
        .unwrap_err();
        assert_eq!(a.failure.code, EditFailureCode::TrackedRevisionConflict);
        assert_eq!(normalized(&a), normalized(&b));
        assert_eq!(direct.direct_batches_applied(), 0);
        assert_eq!(base, direct.encode_state_as_update_v1());
        assert_eq!(base, replica.encode_state_as_update_v1());
        assert_stories(&direct, &replica);
        assert_export(&direct, &replica, &bytes);
    }
}

#[test]
fn surrogate_pairs_and_run_boundaries() {
    let bytes = docx(&format!(
        "{}{}{}{}{}",
        p(
            "00000001",
            &format!(
                "{}{}{}",
                r("A😀"),
                r#"<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>"#,
                r("🦀Z")
            )
        ),
        p("00000002", &r("Delta epsilon")),
        p("00000003", ""),
        p("00000004", &r("User")),
        p("00000005", &r("Probe"))
    ));
    exercise(
        &bytes,
        |_| {},
        vec![
            suggested(replace(range("00000001", 1, 3), "🚀")),
            replace(range("00000001", 5, 9), "BI"),
            suggested(delete(range("00000001", 13, 15))),
        ],
    );
}

#[test]
fn insert_at_start_of_empty_paragraph_and_mixed_noop() {
    let result = exercise(
        &fixture(),
        |_| {},
        vec![
            insert(range("00000003", 0, 0), "empty"),
            replace(search("00000001", "beta"), "beta"),
            suggested(insert(range("00000002", 0, 0), "new ")),
        ],
    );
    assert!(!result.receipts[1].changed);
}

#[test]
fn separate_histories_and_structural_operations_use_replicas() {
    use docx_edit::content_controls::ContentControlSelector;
    let control = r#"<w:sdt><w:sdtPr><w:tag w:val="inline"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt>"#;
    let bytes = docx(&format!(
        "{}{}",
        p("00000001", &r("Alpha beta")),
        p("00000002", control)
    ));
    let cases = [
        (
            EditHistory::None,
            EditStep::new(EditOperation::InsertParagraphs {
                target: ParagraphTarget {
                    story: "body".to_owned(),
                    para_id: "00000001".to_owned(),
                },
                at: TargetEdge::End,
                paragraphs: vec![ParagraphInput {
                    text: "New".to_owned(),
                    style_id: None,
                }],
            }),
        ),
        (
            EditHistory::None,
            EditStep::new(EditOperation::SetParagraphStyle {
                target: ParagraphTarget {
                    story: "body".to_owned(),
                    para_id: "00000001".to_owned(),
                },
                style_id: "Quote".to_owned(),
            }),
        ),
        (
            EditHistory::None,
            EditStep::new(EditOperation::SetContentControlText {
                target: ContentControlSelector::Tag {
                    tag: "inline".to_owned(),
                },
                text: "filled".to_owned(),
            }),
        ),
        (
            EditHistory::Separate,
            replace(search("00000001", "beta"), "BETA"),
        ),
    ];
    for (history, step) in cases {
        let direct = open(&bytes);
        let replica = open(&bytes);
        assert_eq!(
            direct.encode_state_as_update_v1(),
            replica.encode_state_as_update_v1()
        );
        direct.set_direct_batches(true);
        let a = apply_request(
            &direct,
            &request(&direct, vec![step.clone()], history),
            &UndoSession::new(),
        )
        .unwrap();
        let b = apply_request(
            &replica,
            &request(&replica, vec![step], history),
            &UndoSession::new(),
        )
        .unwrap();
        assert!(a.applied);
        assert_eq!(direct.direct_batches_applied(), 0);
        assert_eq!(normalized(&a), normalized(&b));
        assert_stories(&direct, &replica);
    }
}

#[test]
fn refusals_preserve_state_and_match() {
    for stale in [false, true] {
        let direct = open(&fixture());
        let replica = open(&fixture());
        let base = direct.encode_state_as_update_v1();
        assert_eq!(base, replica.encode_state_as_update_v1());
        direct.set_direct_batches(true);
        let steps = if stale {
            vec![replace(search("00000001", "beta"), "BETA")]
        } else {
            vec![
                delete(search("00000001", "beta")),
                replace(search("00000001", "beta"), "BETA"),
            ]
        };
        let a = request(&direct, steps.clone(), EditHistory::None);
        let b = request(&replica, steps, EditHistory::None);
        if stale {
            for doc in [&direct, &replica] {
                doc.begin_opening(Some("changed"));
            }
        }
        let before = direct.encode_state_as_update_v1();
        let updates_a = Updates::new(&direct);
        let updates_b = Updates::new(&replica);
        let a = apply_request(&direct, &a, &UndoSession::new()).unwrap_err();
        let b = apply_request(&replica, &b, &UndoSession::new()).unwrap_err();
        assert_eq!(
            a.failure.code,
            if stale {
                EditFailureCode::StaleVersion
            } else {
                EditFailureCode::OverlappingSteps
            }
        );
        assert_eq!(normalized(&a), normalized(&b));
        assert_eq!(direct.direct_batches_applied(), 0);
        assert_eq!(before, direct.encode_state_as_update_v1());
        assert_eq!(before, replica.encode_state_as_update_v1());
        assert!(updates_a.events.lock().unwrap().is_empty());
        assert!(updates_b.events.lock().unwrap().is_empty());
    }
}

#[test]
fn leading_format_marker_batches_match_the_replica_path() {
    let bytes = docx(&format!(
        "{}{}",
        p(
            "00000001",
            &format!(
                r#"<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">AB</w:t></w:r>{}"#,
                r("X")
            )
        ),
        p("00000002", &r("Tail")),
    ));
    let direct = open(&bytes);
    let replica = open(&bytes);
    direct.set_direct_batches(true);
    let updates_direct = Updates::new(&direct);
    let updates_replica = Updates::new(&replica);
    let (direct_undo, replica_undo) = (UndoSession::new(), UndoSession::new());
    for (steps, applied) in [
        (vec![delete(range("00000001", 0, 2))], 1),
        (vec![insert(range("00000001", 0, 0), "Y")], 2),
    ] {
        let result_direct = apply(&direct, &direct_undo, steps.clone());
        let result_replica = apply(&replica, &replica_undo, steps);
        assert!(result_direct.applied);
        assert_eq!(normalized(&result_direct), normalized(&result_replica));
        assert_eq!(direct.direct_batches_applied(), applied);
        assert_eq!(replica.direct_batches_applied(), 0);
        assert_eq!(
            direct.encode_state_as_update_v1(),
            replica.encode_state_as_update_v1()
        );
        assert_eq!(state_vector(&direct), state_vector(&replica));
        assert_stories(&direct, &replica);
        assert_export(&direct, &replica, &bytes);
    }
    assert_eq!(updates_direct.signature(), updates_replica.signature());
}

#[test]
fn direct_and_replica_peers_converge() {
    let bytes = fixture();
    let seed = open(&bytes).encode_state_as_update_v1();
    let direct = EditingDoc::new(7101);
    let replica = EditingDoc::new(7102);
    for doc in [&direct, &replica] {
        doc.apply_update_v1(&seed).unwrap();
        doc.retain_source_docx(bytes.clone());
    }
    assert_eq!(
        direct.encode_state_as_update_v1(),
        replica.encode_state_as_update_v1()
    );
    direct.set_direct_batches(true);
    let updates_a = Updates::new(&direct);
    let updates_b = Updates::new(&replica);
    apply(
        &direct,
        &UndoSession::new(),
        vec![suggested(replace(search("00000001", "beta"), "BETA"))],
    );
    apply(
        &replica,
        &UndoSession::new(),
        vec![suggested(insert(range("00000002", 0, 0), "New "))],
    );
    assert_eq!(direct.direct_batches_applied(), 1);
    assert_eq!(replica.direct_batches_applied(), 0);
    let a: Vec<_> = updates_a
        .events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect();
    let b: Vec<_> = updates_b
        .events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect();
    for update in b {
        direct.apply_update_v1(&update).unwrap();
    }
    for update in a {
        replica.apply_update_v1(&update).unwrap();
    }
    assert_eq!(state_vector(&direct), state_vector(&replica));
    assert_stories(&direct, &replica);
    assert_export(&direct, &replica, &bytes);
}

#[test]
fn noop_batches_leave_direct_counter_and_events_unchanged() {
    let direct = open(&fixture());
    let replica = open(&fixture());
    let base = direct.encode_state_as_update_v1();
    assert_eq!(base, replica.encode_state_as_update_v1());
    direct.set_direct_batches(true);
    let updates_direct = Updates::new(&direct);
    let updates_replica = Updates::new(&replica);
    let steps = vec![
        replace(search("00000001", "beta"), "beta"),
        insert(range("00000003", 0, 0), ""),
    ];
    let a = apply(&direct, &UndoSession::new(), steps.clone());
    let b = apply(&replica, &UndoSession::new(), steps);
    assert!(!a.applied);
    assert_eq!(normalized(&a), normalized(&b));
    assert_eq!(direct.direct_batches_applied(), 0);
    assert_eq!(base, direct.encode_state_as_update_v1());
    assert_eq!(base, replica.encode_state_as_update_v1());
    assert!(updates_direct.events.lock().unwrap().is_empty());
    assert!(updates_replica.events.lock().unwrap().is_empty());
}
