//! Content-control discovery and `setContentControlText` steps against the synthetic template
//! and small single-feature packages.

#[allow(dead_code)]
#[path = "support/structured_fixture.rs"]
mod fixture;

use std::collections::HashMap;

use docx_edit::content_controls::{
    Anchor, ContentControl, ContentControlQuery, ContentControlsOptions, ContentControlsSnapshot,
    ControlPlacement, ControlValue, DiagnosticCode, StorySelection, ValueUnavailable,
    find_docx_content_controls, list_docx_content_controls, list_package_content_controls,
};
use docx_edit::structured::{
    BlockKind, ExportFailureCode, ExportOptions, InlineKind, RevisionView, export_docx_structured,
};
use docx_edit::{
    EditApplication, EditFailureCode, EditFailureReason, EditRefusal, EditRequest, EditingDoc,
    SegmentContent, UndoSession, seed_from_docx,
};
use fixture::{Package, para, run};
use serde_json::{Value, json};
use yrs::{Any, ReadTxn, Text, Transact};

fn template() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx"
    ))
    .unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(4242);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn list(doc: &EditingDoc) -> ContentControlsSnapshot {
    doc.list_content_controls(&ContentControlsOptions::default())
        .unwrap()
        .content
}

fn by_tag<'a>(snapshot: &'a ContentControlsSnapshot, tag: &str) -> &'a ContentControl {
    snapshot
        .controls
        .iter()
        .find(|control| control.metadata.tag.as_deref() == Some(tag))
        .unwrap_or_else(|| panic!("no control tagged {tag}"))
}

fn text(control: &ContentControl) -> &str {
    match &control.value {
        ControlValue::Text { text } => text,
        other => panic!("no text value: {other:?}"),
    }
}

fn by_id(id: &str, text: &str) -> Value {
    json!({"op": "setContentControlText", "target": {"kind": "id", "controlId": id}, "text": text})
}

fn by_tag_step(tag: &str, text: &str) -> Value {
    json!({"op": "setContentControlText", "target": {"kind": "tag", "tag": tag}, "text": text})
}

fn request(doc: &EditingDoc, steps: Vec<Value>) -> EditRequest {
    serde_json::from_value(json!({"expectVersion": doc.version().as_str(), "steps": steps}))
        .unwrap()
}

fn apply(doc: &EditingDoc, undo: &UndoSession, steps: Vec<Value>) -> EditApplication {
    doc.apply_edits(&request(doc, steps), undo)
        .unwrap()
        .unwrap_or_else(|refusal| panic!("refused: {refusal:?}"))
}

fn refuse(doc: &EditingDoc, steps: Vec<Value>) -> EditRefusal {
    match doc
        .apply_edits(&request(doc, steps), &UndoSession::new())
        .unwrap()
    {
        Ok(applied) => panic!("applied: {applied:?}"),
        Err(refusal) => refusal,
    }
}

fn reason(doc: &EditingDoc, steps: Vec<Value>) -> (EditFailureCode, Option<EditFailureReason>) {
    let failure = refuse(doc, steps).failure;
    (failure.code, failure.reason)
}

fn embed_payload(
    doc: &EditingDoc,
    story: &str,
    tag: &str,
) -> std::collections::BTreeMap<String, Any> {
    doc.story_segments(story)
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::OtherEmbed { payload, .. }
                if payload.get("tag") == Some(&Any::from(tag)) =>
            {
                Some(payload)
            }
            _ => None,
        })
        .unwrap_or_else(|| panic!("no embed tagged {tag}"))
}

fn inline_sdt(properties: &str, content: &str) -> String {
    format!("<w:sdt><w:sdtPr>{properties}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>")
}

fn package(body: &str) -> Vec<u8> {
    Package::new(body).bytes()
}

#[test]
fn lists_the_template_controls_in_document_order() {
    let snapshot =
        list_docx_content_controls(&template(), &ContentControlsOptions::default()).unwrap();
    let summary: Vec<(&str, Option<&str>, ControlPlacement)> = snapshot
        .controls
        .iter()
        .map(|control| {
            (
                control.metadata.control_id.as_str(),
                control.metadata.tag.as_deref(),
                control.placement,
            )
        })
        .collect();
    assert_eq!(
        summary,
        [
            (
                "body|10000002|0",
                Some("customer.name"),
                ControlPlacement::Inline
            ),
            (
                "body|10000003|0",
                Some("account.reference"),
                ControlPlacement::Inline
            ),
            (
                "body|10000004|0",
                Some("account.reference"),
                ControlPlacement::Inline
            ),
            (
                "body|10000005|0",
                Some("terms.standard"),
                ControlPlacement::Inline
            ),
            (
                "body:sdt0",
                Some("customer.address"),
                ControlPlacement::Block
            ),
            (
                "body|10000007|0",
                Some("customer.email"),
                ControlPlacement::Inline
            ),
            (
                "hf:rIdHeader|20000001|0",
                Some("document.title"),
                ControlPlacement::Inline
            ),
        ]
    );
    assert!(snapshot.complete);
    assert!(
        snapshot.diagnostics.is_empty(),
        "{:?}",
        snapshot.diagnostics
    );
    let name = by_tag(&snapshot, "customer.name");
    assert_eq!(name.metadata.ooxml_id.as_deref(), Some("101"));
    assert_eq!(name.metadata.control_type, "plainText");
    assert!(name.metadata.showing_placeholder);
    assert_eq!(text(name), "Click to enter a name.");
    assert_eq!(name.multi_line, Some(false));
    assert_eq!(
        name.anchor,
        Anchor::Control {
            story: "body".to_owned(),
            control_id: "body|10000002|0".to_owned()
        }
    );
    let address = by_tag(&snapshot, "customer.address");
    assert_eq!(text(address), "1 Old Road\nOldtown");
    assert_eq!(address.multi_line, Some(true));
    let reference = by_tag(&snapshot, "account.reference");
    assert_eq!(reference.metadata.control_type, "richText");
    assert_eq!(reference.multi_line, None);
    let locked = by_tag(&snapshot, "terms.standard");
    assert_eq!(locked.metadata.lock.as_deref(), Some("contentLocked"));
    assert!(locked.effective_lock.content && !locked.effective_lock.control);
    assert!(by_tag(&snapshot, "customer.email").metadata.data_bound);
    let wire = serde_json::to_string(name).unwrap();
    let value = serde_json::to_value(name).unwrap();
    let mut keys: Vec<(usize, String)> = value
        .as_object()
        .unwrap()
        .keys()
        .map(|key| (wire.find(&format!("\"{key}\":")).unwrap(), key.clone()))
        .collect();
    keys.sort();
    let keys: Vec<String> = keys.into_iter().map(|(_, key)| key).collect();
    assert_eq!(
        keys,
        [
            "controlId",
            "ooxmlId",
            "controlType",
            "tag",
            "alias",
            "lock",
            "showingPlaceholder",
            "dataBound",
            "placement",
            "anchor",
            "parentControlId",
            "value",
            "multiLine",
            "effectiveLock"
        ]
    );
}

#[test]
fn control_ids_are_the_structured_export_ids() {
    let bytes = template();
    let export = export_docx_structured(
        &bytes,
        &ExportOptions {
            stories: Some(vec![StorySelection::Body, StorySelection::Headers]),
            ..ExportOptions::new(RevisionView::Accepted)
        },
    )
    .unwrap();
    let mut exported = Vec::new();
    for story in &export.stories {
        for block in &story.blocks {
            match &block.content {
                BlockKind::ContentControl { control, .. } => {
                    exported.push(control.control_id.clone())
                }
                BlockKind::Paragraph { paragraph } | BlockKind::Heading { paragraph, .. } => {
                    for inline in &paragraph.inlines {
                        if let InlineKind::ContentControl { control, .. } = &inline.content {
                            exported.push(control.control_id.clone());
                        }
                    }
                }
                _ => {}
            }
        }
    }
    let listed: Vec<String> =
        list_docx_content_controls(&bytes, &ContentControlsOptions::default())
            .unwrap()
            .controls
            .into_iter()
            .map(|control| control.metadata.control_id)
            .collect();
    assert_eq!(listed, exported);
}

#[test]
fn find_matches_exactly_and_returns_every_match() {
    let bytes = template();
    let find = |query: Value| {
        let query: ContentControlQuery = serde_json::from_value(query).unwrap();
        find_docx_content_controls(&bytes, &query, &ContentControlsOptions::default())
            .unwrap()
            .controls
            .into_iter()
            .map(|control| control.metadata.control_id)
            .collect::<Vec<_>>()
    };
    assert_eq!(
        find(json!({"kind": "tag", "tag": "account.reference"})),
        ["body|10000003|0", "body|10000004|0"]
    );
    assert!(find(json!({"kind": "tag", "tag": "Account.Reference"})).is_empty());
    assert!(find(json!({"kind": "tag", "tag": "account.reference "})).is_empty());
    assert_eq!(
        find(json!({"kind": "alias", "alias": "Customer name"})),
        ["body|10000002|0"]
    );
    assert_eq!(
        find(json!({"kind": "id", "controlId": "body:sdt0"})),
        ["body:sdt0"]
    );

    let body = para(
        "20000001",
        &format!(
            "{}{}{}",
            inline_sdt(
                r#"<w:tag w:val=""/><w:id w:val="7"/><w:text/>"#,
                &run("empty tag")
            ),
            inline_sdt(r#"<w:id w:val="7"/><w:text/>"#, &run("no tag")),
            inline_sdt("", &run("bare"))
        ),
    );
    let bytes = package(&body);
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    let tags: Vec<Option<&str>> = snapshot
        .controls
        .iter()
        .map(|control| control.metadata.tag.as_deref())
        .collect();
    assert_eq!(tags, [Some(""), None, None]);
    assert_eq!(snapshot.controls[0].metadata.ooxml_id.as_deref(), Some("7"));
    assert_eq!(snapshot.controls[1].metadata.ooxml_id.as_deref(), Some("7"));
    assert_ne!(
        snapshot.controls[0].metadata.control_id,
        snapshot.controls[1].metadata.control_id
    );
    assert_eq!(snapshot.controls[2].metadata.control_type, "richText");
    let query: ContentControlQuery =
        serde_json::from_value(json!({"kind": "tag", "tag": ""})).unwrap();
    let empty =
        find_docx_content_controls(&bytes, &query, &ContentControlsOptions::default()).unwrap();
    assert_eq!(empty.controls.len(), 1);
}

#[test]
fn selections_and_budgets_refuse_rather_than_truncate() {
    let bytes = template();
    let headers = list_docx_content_controls(
        &bytes,
        &ContentControlsOptions {
            stories: Some(vec![StorySelection::Headers]),
            ..ContentControlsOptions::default()
        },
    )
    .unwrap();
    assert_eq!(headers.controls.len(), 1);
    assert_eq!(headers.included_stories, [StorySelection::Headers]);
    assert!(headers.complete);
    assert!(headers.diagnostics.iter().any(|diagnostic| diagnostic.code
        == DiagnosticCode::StoriesOmitted
        && diagnostic.message.starts_with("6 content controls in body")));
    let refusal =
        |options: ContentControlsOptions| match list_docx_content_controls(&bytes, &options) {
            Err(docx_edit::structured::ExportError::Refused(failure)) => failure.code,
            other => panic!("not refused: {other:?}"),
        };
    assert_eq!(
        refusal(ContentControlsOptions {
            max_controls: Some(3),
            ..ContentControlsOptions::default()
        }),
        ExportFailureCode::LimitExceeded
    );
    assert_eq!(
        refusal(ContentControlsOptions {
            max_bytes: Some(1_024),
            ..ContentControlsOptions::default()
        }),
        ExportFailureCode::LimitExceeded
    );
    assert_eq!(
        refusal(ContentControlsOptions {
            max_controls: Some(0),
            ..ContentControlsOptions::default()
        }),
        ExportFailureCode::InvalidOptions
    );
    assert_eq!(
        refusal(ContentControlsOptions {
            stories: Some(Vec::new()),
            ..ContentControlsOptions::default()
        }),
        ExportFailureCode::InvalidOptions
    );
    let doc = open(&bytes);
    let refused = doc
        .list_content_controls(&ContentControlsOptions {
            max_controls: Some(2),
            ..ContentControlsOptions::default()
        })
        .unwrap_err();
    assert_eq!(refused.version, doc.version());
    assert_eq!(refused.failure.code, ExportFailureCode::LimitExceeded);
}

#[test]
fn nested_controls_are_listed_before_their_descendants_with_inherited_locks() {
    let inner = inline_sdt(r#"<w:tag w:val="inner"/><w:text/>"#, &run("inner"));
    let outer = inline_sdt(
        r#"<w:tag w:val="outer"/><w:lock w:val="contentLocked"/>"#,
        &format!("{}{inner}", run("outer ")),
    );
    let cell = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        para(
            "30000003",
            &inline_sdt(
                r#"<w:tag w:val="cell"/><w:lock w:val="bogus"/>"#,
                &run("cell")
            )
        )
    );
    let block = format!(
        "<w:sdt><w:sdtPr><w:tag w:val=\"block\"/></w:sdtPr><w:sdtContent>{}{cell}</w:sdtContent></w:sdt>",
        para(
            "30000002",
            &inline_sdt(r#"<w:tag w:val="child"/>"#, &run("child"))
        )
    );
    let body = format!(
        "{}{block}{}",
        para("30000001", &outer),
        para("30000004", &run("end"))
    );
    let snapshot =
        list_docx_content_controls(&package(&body), &ContentControlsOptions::default()).unwrap();
    let order: Vec<(&str, Option<&str>)> = snapshot
        .controls
        .iter()
        .map(|control| {
            (
                control.metadata.tag.as_deref().unwrap(),
                control.parent_control_id.as_deref(),
            )
        })
        .collect();
    assert_eq!(
        order,
        [
            ("outer", None),
            ("inner", Some("body|30000001|0")),
            ("block", None),
            ("child", Some("body:sdt0")),
            ("cell", Some("body:sdt0")),
        ]
    );
    assert_eq!(
        snapshot.controls[1].metadata.control_id,
        "body|30000001|0|0"
    );
    assert!(snapshot.controls[1].effective_lock.content);
    assert!(snapshot.controls[1].effective_lock.control);
    assert_eq!(
        snapshot.controls[0].value,
        ControlValue::Unavailable {
            reason: ValueUnavailable::NonTextContent
        }
    );
    assert!(!snapshot.controls[4].effective_lock.known);
    let doc = open(&package(&body));
    assert_eq!(
        reason(&doc, vec![by_tag_step("inner", "x")]),
        (
            EditFailureCode::LockedTarget,
            Some(EditFailureReason::ContentLocked)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("child", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::NestedControls)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("block", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::NestedControls)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("cell", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnknownLock)
        )
    );
}

#[test]
fn fills_two_controls_by_id_as_one_undo_step() {
    let doc = open(&template());
    let undo = UndoSession::new();
    let before = list(&doc);
    let applied = apply(
        &doc,
        &undo,
        vec![
            by_id("body|10000002|0", "Ada Lovelace"),
            by_id("body:sdt0", "12 Example Street\r\nLondon\tUK"),
        ],
    );
    assert!(applied.applied);
    assert_eq!(applied.changed_stories, ["body", "body:sdt0"]);
    assert_eq!(
        applied.receipts[0].control.as_ref().unwrap().control_id,
        "body|10000002|0"
    );
    assert_eq!(applied.receipts[1].new_paragraphs.len(), 1);
    assert_eq!(
        applied.receipts[1].control.as_ref().unwrap().anchor,
        Anchor::Control {
            story: "body".to_owned(),
            control_id: "body:sdt0".to_owned()
        }
    );
    let after = list(&doc);
    assert_eq!(text(by_tag(&after, "customer.name")), "Ada Lovelace");
    assert!(!by_tag(&after, "customer.name").metadata.showing_placeholder);
    assert_eq!(
        text(by_tag(&after, "customer.address")),
        "12 Example Street\nLondon\tUK"
    );
    let payload = embed_payload(&doc, "body", "customer.name");
    let Some(Any::String(raw)) = payload.get("rawPropertiesXml") else {
        panic!("raw properties");
    };
    assert!(!raw.contains("showingPlcHdr"));
    assert!(raw.contains("<w15:appearance w15:val=\"tags\"/>"));
    let Some(Any::String(properties)) = payload.get("propertiesJson") else {
        panic!("properties json");
    };
    let properties: Value = serde_json::from_str(properties).unwrap();
    assert!(properties.get("showingPlaceholder").is_none());
    assert_eq!(
        properties["rawPropertiesXml"],
        Value::String(raw.to_string())
    );
    let Some(Any::Array(content)) = payload.get("content") else {
        panic!("content");
    };
    let Any::Map(first) = &content[0] else {
        panic!("item")
    };
    let Some(Any::Map(attrs)) = first.get("attrs") else {
        panic!("attrs")
    };
    assert_eq!(attrs.get("bold"), Some(&Any::Bool(true)));
    assert!(
        !attrs.contains_key("runStyle"),
        "placeholder formatting is not inherited"
    );
    assert_eq!(doc.paragraphs("body:sdt0").unwrap().len(), 2);

    assert!(undo.undo());
    let undone = list(&doc);
    assert_eq!(
        serde_json::to_value(&undone.controls).unwrap(),
        serde_json::to_value(&before.controls).unwrap()
    );
}

#[test]
fn fills_by_unique_tag_and_keeps_formatting_of_the_first_run() {
    let doc = open(&template());
    let undo = UndoSession::new();
    apply(
        &doc,
        &undo,
        vec![by_tag_step("document.title", "Master services agreement")],
    );
    let snapshot = list(&doc);
    assert_eq!(
        text(by_tag(&snapshot, "document.title")),
        "Master services agreement"
    );

    let applied = apply(&doc, &undo, vec![by_id("body|10000003|0", "A\tB\nC")]);
    assert!(applied.applied);
    let payload = embed_payload(&doc, "body", "account.reference");
    let Some(Any::Array(content)) = payload.get("content") else {
        panic!("content");
    };
    let kinds: Vec<String> = content
        .iter()
        .map(|item| match item {
            Any::Map(item) => item
                .get("kind")
                .map(|kind| kind.to_string())
                .unwrap_or_default(),
            _ => String::new(),
        })
        .collect();
    assert_eq!(
        kinds,
        ["A", "tab", "B", "break", "C"].map(|kind| match kind {
            "tab" | "break" => kind.to_owned(),
            _ => "text".to_owned(),
        })
    );
    for item in content.iter() {
        let Any::Map(item) = item else { panic!("item") };
        let Some(Any::Map(attrs)) = item.get("attrs") else {
            panic!("attrs")
        };
        assert_eq!(attrs.get("bold"), Some(&Any::Bool(true)));
    }
}

#[test]
fn equal_text_is_a_no_op_unless_the_placeholder_shows() {
    let doc = open(&template());
    let undo = UndoSession::new();
    let version = doc.version();
    let unchanged = apply(&doc, &undo, vec![by_tag_step("document.title", "Untitled")]);
    assert!(!unchanged.applied);
    assert!(!unchanged.receipts[0].changed);
    assert_eq!(doc.version(), version);
    assert!(!undo.can_undo());

    let content = embed_payload(&doc, "body", "customer.name")
        .get("content")
        .cloned();
    let cleared = apply(
        &doc,
        &undo,
        vec![by_id("body|10000002|0", "Click to enter a name.")],
    );
    assert!(cleared.applied);
    let payload = embed_payload(&doc, "body", "customer.name");
    assert_eq!(payload.get("content").cloned(), content);
    assert_eq!(payload.get("showingPlaceholder"), Some(&Any::Bool(false)));
}

#[test]
fn shrinks_blocks_and_empties_controls() {
    let doc = open(&template());
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_id("body:sdt0", "one\ntwo\nthree")]);
    let ids: Vec<String> = doc
        .paragraphs("body:sdt0")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.para_id)
        .collect();
    assert_eq!(ids.len(), 3);
    assert_eq!(ids[0], "10000006");
    let applied = apply(&doc, &undo, vec![by_id("body:sdt0", "")]);
    assert_eq!(
        applied.receipts[0]
            .removed_paragraphs
            .iter()
            .map(|paragraph| paragraph.para_id.clone())
            .collect::<Vec<_>>(),
        ids[1..]
    );
    let paragraphs = doc.paragraphs("body:sdt0").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].para_id, "10000006");
    assert_eq!(paragraphs[0].text, "");
    apply(&doc, &undo, vec![by_id("body|10000003|0", "")]);
    let payload = embed_payload(&doc, "body", "account.reference");
    assert_eq!(payload.get("content"), Some(&Any::Array(Vec::new().into())));
    assert_eq!(text(by_tag(&list(&doc), "customer.address")), "");
}

#[test]
fn refusals_are_data_and_change_nothing() {
    let doc = open(&template());
    let state = doc.encode_state_as_update_v1();
    let version = doc.version();
    let cases: Vec<(Vec<Value>, EditFailureCode, Option<EditFailureReason>)> = vec![
        (
            vec![by_id("body|99|0", "x")],
            EditFailureCode::MissingTarget,
            Some(EditFailureReason::MissingControl),
        ),
        (
            vec![by_tag_step("nope", "x")],
            EditFailureCode::MissingTarget,
            Some(EditFailureReason::MissingTag),
        ),
        (
            vec![by_tag_step("account.reference", "x")],
            EditFailureCode::AmbiguousTarget,
            Some(EditFailureReason::AmbiguousTag),
        ),
        (
            vec![by_tag_step("terms.standard", "x")],
            EditFailureCode::LockedTarget,
            Some(EditFailureReason::ContentLocked),
        ),
        (
            vec![by_tag_step("customer.email", "x")],
            EditFailureCode::Unsupported,
            Some(EditFailureReason::BoundControl),
        ),
        (
            vec![by_tag_step("document.title", "a\nb")],
            EditFailureCode::InvalidStep,
            Some(EditFailureReason::MultilineNotAllowed),
        ),
        (
            vec![by_tag_step("document.title", "a\rb")],
            EditFailureCode::InvalidStep,
            Some(EditFailureReason::InvalidText),
        ),
        (
            vec![by_tag_step("document.title", "a\u{1}b")],
            EditFailureCode::InvalidStep,
            Some(EditFailureReason::InvalidText),
        ),
        (
            vec![
                json!({"op": "setContentControlText", "target": {"kind": "tag", "tag": "document.title"}, "text": "x", "suggest": {"author": "A", "date": "2026-01-01T00:00:00Z"}}),
            ],
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedSuggestion),
        ),
        (
            vec![
                json!({"op": "setContentControlText", "target": {"kind": "tag", "tag": "document.title"}, "text": "x", "expect": {"text": "Titled"}}),
            ],
            EditFailureCode::ContentMismatch,
            None,
        ),
        (
            vec![
                by_id("body|10000002|0", "valid"),
                by_tag_step("customer.email", "late"),
            ],
            EditFailureCode::Unsupported,
            Some(EditFailureReason::BoundControl),
        ),
        (
            vec![
                by_id("body|10000002|0", "a"),
                by_tag_step("customer.name", "b"),
            ],
            EditFailureCode::OverlappingSteps,
            None,
        ),
    ];
    for (steps, code, why) in cases {
        let refusal = refuse(&doc, steps.clone());
        assert_eq!(
            (refusal.failure.code, refusal.failure.reason),
            (code, why),
            "{steps:?}: {}",
            refusal.failure.message
        );
        assert_eq!(refusal.version, version);
        assert_eq!(doc.encode_state_as_update_v1(), state);
    }
    let stale: EditRequest = serde_json::from_value(
        json!({"expectVersion": "0-0", "steps": [by_tag_step("document.title", "x")]}),
    )
    .unwrap();
    let refusal = doc
        .apply_edits(&stale, &UndoSession::new())
        .unwrap()
        .unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
    let fresh = open(&template());
    apply(
        &fresh,
        &UndoSession::new(),
        vec![by_id("body:sdt0", "a\nb")],
    );
    apply(&doc, &UndoSession::new(), vec![by_id("body:sdt0", "a\nb")]);
    assert_eq!(
        doc.paragraphs("body:sdt0").unwrap()[1].para_id,
        fresh.paragraphs("body:sdt0").unwrap()[1].para_id
    );
}

#[test]
fn conflicts_reserve_the_control_and_its_paragraph() {
    let doc = open(&template());
    let paragraph = json!({"story": "body", "paraId": "10000002"});
    let whole = json!({"kind": "paragraph", "story": "body", "paraId": "10000002"});
    let atom = json!({"kind": "range", "story": "body", "view": "accepted",
        "start": {"paraId": "10000002", "offset": 10}, "end": {"paraId": "10000002", "offset": 10}});
    for other in [
        json!({"op": "insertText", "target": atom, "at": "start", "text": "x"}),
        json!({"op": "setParagraphStyle", "target": paragraph, "styleId": "Title"}),
    ] {
        let failure = refuse(&doc, vec![by_id("body|10000002|0", "Ada"), other.clone()]).failure;
        assert_eq!(failure.code, EditFailureCode::OverlappingSteps, "{other}");
        assert_eq!(failure.conflicting_step_index, Some(0));
    }
    let undo = UndoSession::new();
    let applied = apply(
        &doc,
        &undo,
        vec![
            by_id("body|10000002|0", "Ada"),
            json!({"op": "insertText", "target": whole, "at": "start", "text": "Dear "}),
            by_id("body|10000003|0", "REF-1"),
        ],
    );
    assert!(applied.applied);
    let snapshot = list(&doc);
    assert_eq!(text(by_tag(&snapshot, "customer.name")), "Ada");
    assert_eq!(text(by_tag(&snapshot, "account.reference")), "REF-1");
}

#[test]
fn validation_previews_the_resolved_control_and_reserves_nothing() {
    let doc = open(&template());
    let state = doc.encode_state_as_update_v1();
    let validation = doc
        .validate_edits(&request(
            &doc,
            vec![by_tag_step("customer.address", "a\nb\nc")],
        ))
        .unwrap()
        .unwrap();
    assert!(validation.would_apply);
    assert_eq!(validation.previews[0].new_paragraph_count, 2);
    assert_eq!(
        validation.previews[0].control.as_ref().unwrap().control_id,
        "body:sdt0"
    );
    assert_eq!(
        serde_json::to_value(&validation.previews[0].target).unwrap(),
        json!({"kind": "contentControl", "selector": {"kind": "id", "controlId": "body:sdt0"}})
    );
    assert_eq!(doc.encode_state_as_update_v1(), state);
}

#[test]
fn unsupported_children_and_revisions_refuse_fills() {
    let bookmark = inline_sdt(
        r#"<w:tag w:val="bookmark"/><w:text/>"#,
        &format!(
            r#"<w:bookmarkStart w:id="1" w:name="kept"/>{}<w:bookmarkEnd w:id="1"/>"#,
            run("marked")
        ),
    );
    let tracked = inline_sdt(
        r#"<w:tag w:val="tracked"/><w:richText/>"#,
        &format!(
            r#"{}<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">{}</w:ins>"#,
            run("a"),
            run("b")
        ),
    );
    let field = inline_sdt(
        r#"<w:tag w:val="field"/><w:richText/>"#,
        r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
    );
    let checkbox = inline_sdt(
        r#"<w:tag w:val="checkbox"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>"#,
        &run("\u{2610}"),
    );
    let body = para("40000001", &format!("{bookmark}{tracked}{field}{checkbox}"));
    let bytes = package(&body);
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    let values: HashMap<&str, &ControlValue> = snapshot
        .controls
        .iter()
        .map(|control| (control.metadata.tag.as_deref().unwrap(), &control.value))
        .collect();
    assert_eq!(
        values["bookmark"],
        &ControlValue::Text {
            text: "marked".to_owned()
        }
    );
    assert_eq!(
        values["tracked"],
        &ControlValue::Unavailable {
            reason: ValueUnavailable::TrackedRevisions
        }
    );
    assert_eq!(
        values["field"],
        &ControlValue::Unavailable {
            reason: ValueUnavailable::NonTextContent
        }
    );
    let doc = open(&bytes);
    assert_eq!(
        reason(&doc, vec![by_tag_step("bookmark", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedChildren)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("tracked", "x")]),
        (EditFailureCode::TrackedRevisionConflict, None)
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("field", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedChildren)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("checkbox", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedControlType)
        )
    );
}

#[test]
fn comment_controls_are_listed_but_not_writable() {
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="0" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:p>{}</w:p></w:comment></w:comments>"#,
        fixture::namespaces(),
        inline_sdt(r#"<w:tag w:val="shared"/><w:text/>"#, &run("in a comment"))
    );
    let body = para(
        "50000001",
        &format!(
            r#"<w:commentRangeStart w:id="0"/>{}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>"#,
            inline_sdt(r#"<w:tag w:val="shared"/><w:text/>"#, &run("body"))
        ),
    );
    let bytes = Package::new(&body)
        .part(
            "comments.xml",
            "rIdComments",
            "comments",
            "comments",
            &comments,
        )
        .bytes();
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    assert_eq!(snapshot.controls.len(), 2);
    let comment = &snapshot.controls[1];
    assert!(
        matches!(&comment.anchor, Anchor::SourcePart { part, .. } if part == "word/comments.xml")
    );
    assert_eq!(
        comment.value,
        ControlValue::Unavailable {
            reason: ValueUnavailable::UnsupportedStory
        }
    );
    assert!(snapshot.complete);
    let body_only = list_docx_content_controls(
        &bytes,
        &ContentControlsOptions {
            stories: Some(vec![StorySelection::Body]),
            ..ContentControlsOptions::default()
        },
    )
    .unwrap();
    assert_eq!(body_only.controls.len(), 1);
    let doc = open(&bytes);
    assert_eq!(
        reason(&doc, vec![by_tag_step("shared", "x")]),
        (
            EditFailureCode::AmbiguousTarget,
            Some(EditFailureReason::AmbiguousTag)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_id(&comment.metadata.control_id, "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedStory)
        )
    );
}

#[test]
fn sessions_without_source_list_but_do_not_fill() {
    let seeded = open(&template());
    let doc = EditingDoc::new(77);
    doc.apply_update_v1(&seeded.encode_state_as_update_v1())
        .unwrap();
    let snapshot = list(&doc);
    assert!(!snapshot.complete);
    assert!(
        snapshot
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == DiagnosticCode::ProvenanceUnavailable)
    );
    assert_eq!(
        snapshot
            .controls
            .iter()
            .filter(|control| control.metadata.control_id.starts_with("body"))
            .count(),
        6
    );
    assert_eq!(
        reason(&doc, vec![by_id("body|10000003|0", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::ProvenanceUnavailable)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("customer.name", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::ProvenanceUnavailable)
        )
    );
}

#[test]
fn reads_do_not_mutate_and_report_the_session_version() {
    let doc = open(&template());
    let state = doc.encode_state_as_update_v1();
    let read = doc
        .list_content_controls(&ContentControlsOptions::default())
        .unwrap();
    assert_eq!(read.version, doc.version());
    assert_eq!(
        serde_json::to_value(read.content.anchor_scope).unwrap(),
        json!("session")
    );
    let query: ContentControlQuery =
        serde_json::from_value(json!({"kind": "tag", "tag": "customer.name"})).unwrap();
    let found = doc
        .find_content_controls(&query, &ContentControlsOptions::default())
        .unwrap();
    assert_eq!(found.content.controls.len(), 1);
    assert_eq!(doc.encode_state_as_update_v1(), state);
    let txn = doc.yrs_doc().transact();
    assert!(txn.store().pending_update().is_none());
}

#[test]
fn text_control_values_cannot_be_written_through_embed_paths() {
    let doc = open(&template());
    let at = doc
        .story_segments("body")
        .unwrap()
        .iter()
        .scan(0u32, |offset, segment| {
            let start = *offset;
            *offset += match &segment.content {
                SegmentContent::Text(text) => text.encode_utf16().count() as u32,
                _ => 1,
            };
            Some((start, segment.clone()))
        })
        .find_map(|(start, segment)| match segment.content {
            SegmentContent::OtherEmbed { payload, .. }
                if payload.get("tag") == Some(&Any::from("document.title"))
                    || payload.get("tag") == Some(&Any::from("customer.name")) =>
            {
                Some(start)
            }
            _ => None,
        })
        .unwrap();
    let ctx = docx_edit::EditCtx::local("", "");
    assert!(matches!(
        doc.set_embed_attrs(
            &ctx,
            docx_edit::Position::new("body", at),
            vec![("value".to_owned(), Any::from("legacy"))]
        ),
        Err(docx_edit::OpError::TextControlValue)
    ));
    assert!(matches!(
        doc.set_embed_attrs_by_id(&ctx, "101", vec![("value".to_owned(), Any::from("legacy"))]),
        Err(docx_edit::OpError::TextControlValue)
    ));
    assert!(matches!(
        doc.apply_raw_ops(
            "body",
            vec![docx_edit::RawOp::SetEmbedAttr {
                index: at,
                key: "value".to_owned(),
                value: Any::from("legacy")
            }],
            &ctx
        ),
        Err(docx_edit::OpError::TextControlValue)
    ));
    assert!(matches!(
        doc.insert_embed(
            &ctx,
            docx_edit::Position::new("body", 0),
            "sdt",
            vec![
                ("sdtType".to_owned(), Any::from("plainText")),
                ("value".to_owned(), Any::from("legacy"))
            ]
        ),
        Err(docx_edit::OpError::TextControlValue)
    ));
    doc.set_embed_attrs_by_id(&ctx, "101", vec![("value".to_owned(), Any::Null)])
        .unwrap();
    doc.insert_embed(
        &ctx,
        docx_edit::Position::new("body", 0),
        "sdt",
        vec![
            ("sdtType".to_owned(), Any::from("checkbox")),
            ("value".to_owned(), Any::from("kept")),
        ],
    )
    .unwrap();
}

#[test]
fn values_updates_give_text_controls_are_read_past_and_filled_away() {
    let doc = open(&template());
    let peer = EditingDoc::new(9);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    let before = peer.encode_state_vector_v1();
    {
        let stories = peer.yrs_doc().transact().get_map("stories").unwrap();
        let mut txn = peer.yrs_doc().transact_mut();
        let Some(yrs::Out::YText(body)) = yrs::Map::get(&stories, &txn, "body") else {
            panic!("body");
        };
        let map = body
            .diff(&txn, yrs::types::text::YChange::identity)
            .into_iter()
            .find_map(|diff| match diff.insert {
                yrs::Out::YMap(map)
                    if yrs::Map::get(&map, &txn, "tag")
                        == Some(yrs::Out::Any(Any::from("customer.name"))) =>
                {
                    Some(map)
                }
                _ => None,
            })
            .unwrap();
        yrs::Map::insert(&map, &mut txn, "value", "legacy");
    }
    doc.apply_update_v1(&peer.encode_diff_v1(&doc.encode_state_vector_v1()).unwrap())
        .unwrap();
    let snapshot = list(&doc);
    assert_eq!(
        text(by_tag(&snapshot, "customer.name")),
        "Click to enter a name."
    );
    assert!(
        snapshot
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == DiagnosticCode::LegacyControlValue)
    );
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_tag_step("customer.name", "Ada")]);
    peer.apply_update_v1(&doc.encode_diff_v1(&before).unwrap())
        .unwrap();
    for replica in [&doc, &peer] {
        assert_eq!(
            embed_payload(replica, "body", "customer.name").get("value"),
            None
        );
    }
    assert!(
        list(&doc)
            .diagnostics
            .iter()
            .all(|diagnostic| diagnostic.code != DiagnosticCode::LegacyControlValue)
    );
}

#[test]
fn controls_inside_raw_blocks_are_listed_in_place() {
    let raw = format!(
        r#"<bofx:block>{}</bofx:block>"#,
        inline_sdt(r#"<w:tag w:val="opaque"/><w:text/>"#, &run("kept as XML"))
    );
    let body = format!(
        "{}{raw}{}",
        para(
            "60000001",
            &inline_sdt(r#"<w:tag w:val="before"/>"#, &run("a"))
        ),
        para(
            "60000002",
            &inline_sdt(r#"<w:tag w:val="after"/>"#, &run("b"))
        )
    );
    let bytes = package(&body);
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    let tags: Vec<&str> = snapshot
        .controls
        .iter()
        .map(|control| control.metadata.tag.as_deref().unwrap())
        .collect();
    assert_eq!(tags, ["before", "opaque", "after"]);
    let opaque = &snapshot.controls[1];
    assert!(matches!(
        &opaque.anchor,
        Anchor::SourcePart { part, path, .. } if part == "word/document.xml" && path.len() == 3
    ));
    assert_eq!(
        opaque.value,
        ControlValue::Unavailable {
            reason: ValueUnavailable::UnsupportedStory
        }
    );
    assert!(snapshot.complete);
    let doc = open(&bytes);
    assert_eq!(
        reason(&doc, vec![by_tag_step("opaque", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedStory)
        )
    );
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_tag_step("after", "c")]);
}

#[test]
fn package_listings_read_raw_blocks_from_their_current_content() {
    let raw = format!(
        r#"<bofx:block>{}</bofx:block>"#,
        inline_sdt(r#"<w:tag w:val="opaque"/><w:text/>"#, &run("kept as XML"))
    );
    let bytes = package(&format!(
        "{raw}{}",
        para(
            "60000003",
            &inline_sdt(r#"<w:tag w:val="after"/>"#, &run("b"))
        )
    ));
    let (mut envelope, parts) = docx_parse::parse_docx_s9_wire_parts_with_limits(
        &bytes,
        docx_parse::S9ParseOptions::default(),
        &docx_parse::ParseLimits::default(),
    )
    .unwrap();
    let options = ContentControlsOptions::default();
    let tags = |snapshot: &ContentControlsSnapshot| -> Vec<String> {
        snapshot
            .controls
            .iter()
            .filter_map(|control| control.metadata.tag.clone())
            .collect()
    };
    let unchanged = list_package_content_controls(envelope.clone(), &parts, &options).unwrap();
    assert_eq!(tags(&unchanged), ["opaque", "after"]);
    assert!(unchanged.complete);
    let content = &mut envelope.document.package.document.content;
    let Some(docx_parse::BlockContent::RawXml(block)) = content.first_mut() else {
        panic!("the first block is kept as raw XML");
    };
    let block = std::sync::Arc::make_mut(block);
    block.xml = block.xml.replace("opaque", "changed");
    let changed = list_package_content_controls(envelope, &parts, &options).unwrap();
    assert_eq!(tags(&changed), ["changed", "after"]);
    assert!(changed.complete);
    assert!(matches!(
        &changed.controls[0].anchor,
        Anchor::Control { story, .. } if story == "body"
    ));
    assert!(matches!(
        &unchanged.controls[0].anchor,
        Anchor::SourcePart { part, .. } if part == "word/document.xml"
    ));
}

#[test]
fn controls_held_only_inside_text_boxes_block_tag_writes() {
    let shared = |text: &str| inline_sdt(r#"<w:tag w:val="shared"/><w:text/>"#, &run(text));
    let text_box = format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="9" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>{}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
        para("60000011", &shared("in the box"))
    );
    let bytes = package(&format!(
        "{}{}",
        para("60000010", &text_box),
        para("60000012", &shared("body"))
    ));
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    assert_eq!(snapshot.controls.len(), 1);
    assert!(!snapshot.complete);
    let doc = open(&bytes);
    assert_eq!(list(&doc).controls, snapshot.controls);
    assert_eq!(
        reason(&doc, vec![by_tag_step("shared", "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::ProvenanceUnavailable)
        )
    );
    let undo = UndoSession::new();
    apply(
        &doc,
        &undo,
        vec![by_id(&snapshot.controls[0].metadata.control_id, "filled")],
    );
    assert_eq!(text(by_tag(&list(&doc), "shared")), "filled");
}

fn shared_header() -> Vec<u8> {
    let header = format!(
        r#"<w:hdr {}>{}</w:hdr>"#,
        fixture::namespaces(),
        para(
            "0D000001",
            &inline_sdt(r#"<w:tag w:val="header.title"/><w:text/>"#, &run("Draft"))
        )
    );
    let body = format!(
        r#"<w:p w14:paraId="0D000002"><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdA"/></w:sectPr></w:pPr>{}</w:p>{}<w:sectPr><w:headerReference w:type="default" r:id="rIdB"/></w:sectPr>"#,
        run("One"),
        para("0D000003", &run("Two"))
    );
    Package::new(&body)
        .part("header1.xml", "rIdA", "header", "header", &header)
        .rel("rIdB", "header", "header1.xml")
        .bytes()
}

#[test]
fn a_header_part_read_twice_holds_one_control_and_fills_write_both_copies() {
    let doc = open(&shared_header());
    let snapshot = list(&doc);
    assert_eq!(snapshot.controls.len(), 1);
    assert!(snapshot.complete);
    let undo = UndoSession::new();
    let applied = apply(&doc, &undo, vec![by_tag_step("header.title", "Final")]);
    assert_eq!(applied.receipts.len(), 1);
    assert_eq!(
        applied.changed_stories,
        ["hf:rIdA".to_owned(), "hf:rIdB".to_owned()]
    );
    for story in ["hf:rIdA", "hf:rIdB"] {
        let Some(Any::Array(content)) = embed_payload(&doc, story, "header.title")
            .get("content")
            .cloned()
        else {
            panic!("content");
        };
        assert!(matches!(
            &content[0],
            Any::Map(item) if item.get("text") == Some(&Any::from("Final"))
        ));
    }
    let snapshot = list(&doc);
    assert!(snapshot.complete);
    assert_eq!(text(by_tag(&snapshot, "header.title")), "Final");
}

#[test]
fn diverged_copies_of_a_header_part_refuse_fills() {
    let doc = open(&shared_header());
    let id = list(&doc).controls[0].metadata.control_id.clone();
    doc.delete_range(
        &docx_edit::EditCtx::local("", ""),
        docx_edit::StoryRange::new("hf:rIdB", 0, 1),
    )
    .unwrap();
    let snapshot = list(&doc);
    assert!(!snapshot.complete);
    assert!(snapshot.diagnostics.iter().any(|diagnostic| diagnostic.code
        == DiagnosticCode::AmbiguousIdentity
        && diagnostic.message.contains("different content controls")));
    assert_eq!(
        reason(&doc, vec![by_id(&id, "Final")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedStory)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("header.title", "Final")]).1,
        Some(EditFailureReason::ProvenanceUnavailable)
    );
}

fn block_sdt(tag: &str, content: &str) -> String {
    format!(
        r#"<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:richText/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"#
    )
}

#[test]
fn paragraph_limits_count_every_fill_of_a_batch() {
    let bytes = package(&format!(
        "{}{}",
        block_sdt("first", &para("0E000001", &run("a"))),
        block_sdt("second", &para("0E000002", &run("b")))
    ));
    let lines = vec!["line"; 600].join("\n");
    let doc = open(&bytes);
    assert_eq!(
        reason(
            &doc,
            vec![by_id("body:sdt0", &lines), by_id("body:sdt1", &lines)]
        ),
        (EditFailureCode::LimitExceeded, None)
    );
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_id("body:sdt0", &lines)]);
}

#[test]
fn empty_surviving_paragraphs_take_their_own_formatting() {
    let bold = r#"<w:r><w:rPr><w:b/></w:rPr><w:t>first</w:t></w:r>"#;
    let bytes = package(&block_sdt(
        "notes",
        &format!("{}{}", para("0E000003", bold), para("0E000004", "")),
    ));
    let doc = open(&bytes);
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_id("body:sdt0", "one\ntwo")]);
    let bold_of = |wanted: &str| {
        doc.story_segments("body:sdt0")
            .unwrap()
            .into_iter()
            .find(
                |segment| matches!(&segment.content, SegmentContent::Text(text) if text == wanted),
            )
            .unwrap_or_else(|| panic!("no text {wanted}"))
            .attributes
            .get("bold")
            .cloned()
    };
    assert_eq!(bold_of("one"), Some(Any::Bool(true)));
    assert_eq!(bold_of("two"), None);
}

#[test]
fn locks_are_read_by_namespace_whatever_the_prefix() {
    let alternate = |tag: &str, lock: &str| {
        format!(
            r#"<q:sdt xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:sdtPr><q:tag q:val="{tag}"/>{lock}<q:text/></q:sdtPr><q:sdtContent>{}</q:sdtContent></q:sdt>"#,
            run("x")
        )
    };
    let bytes = package(&format!(
        "{}{}",
        para(
            "0F000001",
            &alternate("alternate", r#"<q:lock q:val="contentLocked"/>"#)
        ),
        para(
            "0F000002",
            &alternate("mixed", r#"<w:lock q:val="contentLocked"/>"#)
        )
    ));
    let doc = open(&bytes);
    let snapshot = list(&doc);
    let alternate = by_tag(&snapshot, "alternate");
    assert_eq!(alternate.metadata.lock.as_deref(), Some("contentLocked"));
    assert!(alternate.effective_lock.content);
    assert_eq!(
        reason(&doc, vec![by_tag_step("alternate", "y")]),
        (
            EditFailureCode::LockedTarget,
            Some(EditFailureReason::ContentLocked)
        )
    );
    assert_eq!(
        reason(&doc, vec![by_tag_step("mixed", "y")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnknownLock)
        )
    );
}

#[test]
fn controls_locked_against_deletion_are_filled() {
    let bytes = package(&para(
        "0F000003",
        &inline_sdt(
            r#"<w:tag w:val="kept"/><w:lock w:val="sdtLocked"/><w:text/>"#,
            &run("old"),
        ),
    ));
    let doc = open(&bytes);
    let undo = UndoSession::new();
    assert!(apply(&doc, &undo, vec![by_tag_step("kept", "new")]).applied);
    let snapshot = list(&doc);
    let kept = by_tag(&snapshot, "kept");
    assert_eq!(text(kept), "new");
    assert_eq!(kept.metadata.lock.as_deref(), Some("sdtLocked"));
    assert!(kept.effective_lock.control && !kept.effective_lock.content);
}

#[test]
fn controls_parsing_leaves_out_block_tag_writes() {
    for wrapper in ["ins", "moveTo"] {
        let tagged = |id: &str, text: &str| {
            inline_sdt(
                &format!(r#"<w:tag w:val="customer.name"/><w:id w:val="{id}"/><w:text/>"#),
                &run(text),
            )
        };
        let bytes = package(&format!(
            r#"{}<w:p w14:paraId="0F000011"><w:{wrapper} w:id="7" w:author="Ada" w:date="2026-01-01T00:00:00Z">{}</w:{wrapper}></w:p>"#,
            para("0F000010", &tagged("1", "kept")),
            tagged("2", "tracked")
        ));
        let snapshot =
            list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
        assert_eq!(snapshot.controls.len(), 1, "{wrapper}");
        assert!(!snapshot.complete, "{wrapper}");
        assert!(
            snapshot.diagnostics.iter().any(|diagnostic| {
                diagnostic.code == DiagnosticCode::ProvenanceUnavailable
                    && matches!(
                        &diagnostic.anchor,
                        Some(Anchor::SourcePart { part, path, .. })
                            if part == "word/document.xml" && path == &[0, 1, 0, 0]
                    )
            }),
            "{wrapper}: {:?}",
            snapshot.diagnostics
        );
        let doc = open(&bytes);
        assert_eq!(list(&doc).controls, snapshot.controls, "{wrapper}");
        assert_eq!(
            reason(&doc, vec![by_tag_step("customer.name", "x")]),
            (
                EditFailureCode::Unsupported,
                Some(EditFailureReason::ProvenanceUnavailable)
            ),
            "{wrapper}"
        );
        let undo = UndoSession::new();
        apply(
            &doc,
            &undo,
            vec![by_id(&snapshot.controls[0].metadata.control_id, "filled")],
        );
        assert_eq!(text(by_tag(&list(&doc), "customer.name")), "filled");
    }
}

fn notes_package(second_footnote_tag: &str) -> Vec<u8> {
    let control = |tag: &str, text: &str| {
        inline_sdt(&format!(r#"<w:tag w:val="{tag}"/><w:text/>"#), &run(text))
    };
    let notes = |root: &str, note: &str, ids: [&str; 2], second: String| {
        format!(
            r#"<w:{root} {}><w:{note} w:id="1">{}</w:{note}><w:{note} w:id="2">{}</w:{note}></w:{root}>"#,
            fixture::namespaces(),
            para(ids[0], &run("plain note")),
            para(ids[1], &second)
        )
    };
    let body = format!(
        "{}{}",
        para("0F000020", &control("body.tag", "body")),
        para(
            "0F000021",
            r#"<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="2"/></w:r>"#
        )
    );
    Package::new(&body)
        .part(
            "footnotes.xml",
            "rIdFootnotes",
            "footnotes",
            "footnotes",
            &notes(
                "footnotes",
                "footnote",
                ["0F000022", "0F000023"],
                control(second_footnote_tag, "in a footnote"),
            ),
        )
        .part(
            "endnotes.xml",
            "rIdEndnotes",
            "endnotes",
            "endnotes",
            &notes(
                "endnotes",
                "endnote",
                ["0F000024", "0F000025"],
                control("note.en", "in an endnote"),
            ),
        )
        .bytes()
}

#[test]
fn controls_in_later_notes_are_represented() {
    let bytes = notes_package("note.fn");
    let snapshot = list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    let tags: Vec<&str> = snapshot
        .controls
        .iter()
        .filter_map(|control| control.metadata.tag.as_deref())
        .collect();
    assert_eq!(tags, ["body.tag", "note.fn", "note.en"]);
    assert!(snapshot.complete, "{:?}", snapshot.diagnostics);
    assert!(
        snapshot
            .diagnostics
            .iter()
            .all(|diagnostic| diagnostic.code != DiagnosticCode::ProvenanceUnavailable)
    );
    let doc = open(&bytes);
    assert!(list(&doc).complete);
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![by_tag_step("body.tag", "filled")]);
    assert_eq!(text(by_tag(&list(&doc), "body.tag")), "filled");
}

#[test]
fn a_duplicate_tag_in_a_later_note_is_ambiguous() {
    let doc = open(&notes_package("body.tag"));
    assert!(list(&doc).complete);
    assert_eq!(
        reason(&doc, vec![by_tag_step("body.tag", "x")]),
        (
            EditFailureCode::AmbiguousTarget,
            Some(EditFailureReason::AmbiguousTag)
        )
    );
}
