//! Authored values on text content controls: collaboration integrates whatever values updates
//! carry, local authoring never introduces, replaces or moves one, retyping drops one, and fills
//! drop them outside undo history.

use docx_edit::content_controls::{
    ContentControlsOptions, ContentControlsSnapshot, ControlValue, DiagnosticCode,
};
use docx_edit::{
    EditCtx, EditRequest, EditingDoc, OpError, Position, RawOp, SegmentContent, StoryRange,
    UndoSession, seed_from_docx, story_checksum,
};
use serde_json::json;
use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, Transact};

const LEGACY_STATE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../packages/docx/src/yrs/__fixtures__/content-controls/legacy-value.bin"
);

fn template() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx"
    ))
    .unwrap()
}

fn seeded(client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    seed_from_docx(&doc, &template()).unwrap();
    doc
}

fn peer(client: u64, state: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(client);
    doc.apply_update_v1(state).unwrap();
    doc
}

/// The embed map of the control tagged `tag` in `story`, written without the session's guards,
/// the way an older version stored an authored value.
fn with_control<R>(
    doc: &EditingDoc,
    story: &str,
    tag: &str,
    write: impl FnOnce(&MapRef, &mut yrs::TransactionMut<'_>) -> R,
) -> R {
    let stories = doc.yrs_doc().transact().get_map("stories").unwrap();
    let mut txn = doc.yrs_doc().transact_mut();
    let Some(Out::YText(text)) = stories.get(&txn, story) else {
        panic!("no story {story}");
    };
    let map = text
        .diff(&txn, YChange::identity)
        .into_iter()
        .find_map(|diff| match diff.insert {
            Out::YMap(map) if map.get(&txn, "tag") == Some(Out::Any(Any::from(tag))) => Some(map),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no control tagged {tag}"));
    write(&map, &mut txn)
}

fn legacy(doc: &EditingDoc, tag: &str, value: &str) {
    with_control(doc, "body", tag, |map, txn| {
        map.insert(txn, "value", value);
    });
}

fn value(doc: &EditingDoc, tag: &str) -> Option<Any> {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::OtherEmbed { payload, .. }
                if payload.get("tag") == Some(&Any::from(tag)) =>
            {
                Some(payload.get("value").cloned())
            }
            _ => None,
        })
        .unwrap()
}

fn listed(doc: &EditingDoc) -> ContentControlsSnapshot {
    doc.list_content_controls(&ContentControlsOptions::default())
        .unwrap()
        .content
}

fn flagged(doc: &EditingDoc) -> bool {
    listed(doc)
        .diagnostics
        .iter()
        .any(|diagnostic| diagnostic.code == DiagnosticCode::LegacyControlValue)
}

fn text_of(doc: &EditingDoc, tag: &str) -> String {
    listed(doc)
        .controls
        .into_iter()
        .find(|control| control.metadata.tag.as_deref() == Some(tag))
        .map(|control| match control.value {
            ControlValue::Text { text } => text,
            other => panic!("{other:?}"),
        })
        .unwrap()
}

/// What `peer` changed since `base`, as the update another replica would receive.
fn since(peer: &EditingDoc, base: &[u8]) -> Vec<u8> {
    peer.encode_diff_v1(base).unwrap()
}

fn sync(from: &EditingDoc, to: &EditingDoc) {
    to.apply_update_v1(&since(from, &to.encode_state_vector_v1()))
        .unwrap();
}

/// Two replicas hold the same state and read and project it the same way.
fn converged(left: &EditingDoc, right: &EditingDoc) {
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(left.story_segments("body"), right.story_segments("body"));
    assert_eq!(
        story_checksum(left, "body").unwrap(),
        story_checksum(right, "body").unwrap()
    );
    assert_eq!(listed(left), listed(right));
}

/// The story index of the control embed tagged `tag` in the body.
fn embed_index(doc: &EditingDoc, tag: &str) -> u32 {
    let mut index = 0;
    for segment in doc.story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => index += text.encode_utf16().count() as u32,
            SegmentContent::OtherEmbed { payload, .. }
                if payload.get("tag") == Some(&Any::from(tag)) =>
            {
                return index;
            }
            _ => index += 1,
        }
    }
    panic!("no control tagged {tag}");
}

fn fill(doc: &EditingDoc, undo: &UndoSession, id: &str, text: &str) -> bool {
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": doc.version().as_str(),
        "steps": [{"op": "setContentControlText", "target": {"kind": "id", "controlId": id}, "text": text}]
    }))
    .unwrap();
    doc.apply_edits(&request, undo)
        .unwrap()
        .unwrap_or_else(|refusal| panic!("{refusal:?}"))
        .applied
}

fn checkbox_value(checked: bool) -> Any {
    Any::from_json(&format!(r#"{{"kind":"checkbox","checked":{checked}}}"#)).unwrap()
}

#[test]
fn peers_converge_through_deleting_restoring_and_redoing_a_valued_control() {
    let author = seeded(11);
    legacy(&author, "account.reference", "REF-LEGACY");
    let state = author.encode_state_as_update_v1();
    let left = peer(21, &state);
    let right = peer(22, &state);
    for replica in [&left, &right] {
        assert_eq!(
            replica.encode_state_vector_v1(),
            author.encode_state_vector_v1()
        );
        assert_eq!(text_of(replica, "account.reference"), "REF-000");
        assert!(flagged(replica));
    }
    let undo = UndoSession::new();
    undo.track(&left);
    let at = embed_index(&left, "account.reference");
    left.delete_range(&EditCtx::local("", ""), StoryRange::new("body", at, at + 1))
        .unwrap();
    sync(&left, &right);
    converged(&left, &right);
    assert!(undo.undo());
    assert_eq!(
        value(&left, "account.reference"),
        Some(Any::from("REF-LEGACY"))
    );
    sync(&left, &right);
    converged(&left, &right);
    assert_eq!(text_of(&right, "account.reference"), "REF-000");
    assert!(undo.redo());
    sync(&left, &right);
    converged(&left, &right);
    sync(&right, &left);
    converged(&left, &right);
}

#[test]
fn remote_updates_integrate_every_value_transition() {
    let author = seeded(12);
    legacy(&author, "account.reference", "REF-LEGACY");
    let state = author.encode_state_as_update_v1();
    let local = peer(31, &state);
    let writer = peer(32, &state);
    legacy(&writer, "customer.name", "Ada");
    sync(&writer, &local);
    legacy(&writer, "account.reference", "REF-OTHER");
    sync(&writer, &local);
    with_control(&writer, "body", "account.reference", |map, txn| {
        map.remove(txn, "value");
    });
    legacy(&writer, "customer.address", "REF-OTHER");
    sync(&writer, &local);
    with_control(&writer, "body", "account.reference", |map, txn| {
        map.insert(txn, "value", checkbox_value(true));
        map.insert(
            txn,
            "content",
            Any::from_json(
                r#"[{"kind":"sdt","attrs":{},"payload":{"sdtType":"plainText","tag":"inner","value":"x","content":[]}}]"#,
            )
            .unwrap(),
        );
    });
    sync(&writer, &local);
    converged(&local, &writer);
    assert_eq!(value(&local, "customer.name"), Some(Any::from("Ada")));
    assert_eq!(
        text_of(&local, "customer.name"),
        text_of(&seeded(33), "customer.name")
    );
    assert_eq!(text_of(&local, "inner"), "");
    assert!(flagged(&local));
}

#[test]
fn updates_integrate_once_their_dependencies_arrive_in_any_order() {
    let doc = seeded(15);
    let remote = peer(61, &doc.encode_state_as_update_v1());
    let base = doc.encode_state_vector_v1();
    let ctx = EditCtx::local("", "");
    remote
        .insert_text(
            &ctx,
            Position::new("body", 0),
            "one ",
            docx_edit::FormatPolicy::Plain,
        )
        .unwrap();
    let middle = remote.encode_state_vector_v1();
    let first = since(&remote, &base);
    remote
        .insert_text(
            &ctx,
            Position::new("body", 4),
            "two ",
            docx_edit::FormatPolicy::Plain,
        )
        .unwrap();
    legacy(&remote, "account.reference", "REF-LATE");
    let second = since(&remote, &middle);
    doc.apply_update_v1(&second).unwrap();
    assert_eq!(value(&doc, "account.reference"), None);
    doc.apply_update_v1(&first).unwrap();
    assert_eq!(
        doc.encode_state_vector_v1(),
        remote.encode_state_vector_v1()
    );
    assert_eq!(doc.story_segments("body"), remote.story_segments("body"));
    assert!(
        doc.paragraphs("body").unwrap()[0]
            .text
            .starts_with("one two ")
    );
    assert_eq!(
        value(&doc, "account.reference"),
        Some(Any::from("REF-LATE"))
    );
}

#[test]
fn local_authoring_never_introduces_replaces_or_moves_a_text_value() {
    let doc = seeded(13);
    legacy(&doc, "account.reference", "REF-LEGACY");
    let ctx = EditCtx::local("", "");
    let name = Position::new("body", embed_index(&doc, "customer.name"));
    let reference = embed_index(&doc, "account.reference");
    fn refused<T: std::fmt::Debug>(result: Result<T, OpError>) {
        assert!(
            matches!(result, Err(OpError::TextControlValue)),
            "{result:?}"
        );
    }
    refused(doc.set_embed_attrs(
        &ctx,
        name.clone(),
        vec![("value".to_owned(), Any::from("Ada"))],
    ));
    refused(doc.set_embed_attrs(
        &ctx,
        name.clone(),
        vec![("value".to_owned(), checkbox_value(true))],
    ));
    refused(doc.set_embed_attrs(
        &ctx,
        Position::new("body", reference),
        vec![("value".to_owned(), Any::from("REF-OTHER"))],
    ));
    refused(doc.apply_raw_ops(
        "body",
        vec![RawOp::SetEmbedAttr {
            index: reference,
            key: "value".to_owned(),
            value: Any::from("REF-OTHER"),
        }],
        &ctx,
    ));
    refused(doc.insert_embed(
        &ctx,
        Position::new("body", 0),
        "sdt",
        vec![
            ("tag".to_owned(), Any::from("copy")),
            ("value".to_owned(), Any::from("REF-LEGACY")),
        ],
    ));
    let nested = Any::from_json(
        r#"[{"kind":"sdt","attrs":{},"payload":{"sdtType":"plainText","tag":"inner","value":"x","content":[]}}]"#,
    )
    .unwrap();
    refused(doc.insert_embed(
        &ctx,
        Position::new("body", 0),
        "sdt",
        vec![("content".to_owned(), nested.clone())],
    ));
    refused(doc.set_embed_attrs(&ctx, name.clone(), vec![("content".to_owned(), nested)]));
    assert_eq!(
        value(&doc, "account.reference"),
        Some(Any::from("REF-LEGACY"))
    );
    assert_eq!(value(&doc, "customer.name"), None);
    doc.set_embed_attrs(
        &ctx,
        Position::new("body", reference),
        vec![("alias".to_owned(), Any::from("Renamed"))],
    )
    .unwrap();
    doc.set_embed_attrs(
        &ctx,
        Position::new("body", reference),
        vec![("value".to_owned(), Any::Null)],
    )
    .unwrap();
    assert_eq!(value(&doc, "account.reference"), None);
}

#[test]
fn typed_setters_write_values_and_retyping_drops_them() {
    let doc = seeded(14);
    let ctx = EditCtx::local("", "");
    doc.insert_embed(
        &ctx,
        Position::new("body", 0),
        "sdt",
        vec![
            ("tag".to_owned(), Any::from("agree")),
            ("sdtType".to_owned(), Any::from("checkbox")),
        ],
    )
    .unwrap();
    doc.set_embed_attrs(
        &ctx,
        Position::new("body", 0),
        vec![("value".to_owned(), checkbox_value(true))],
    )
    .unwrap();
    assert_eq!(value(&doc, "agree"), Some(checkbox_value(true)));
    assert!(!flagged(&doc));
    doc.apply_raw_ops(
        "body",
        vec![RawOp::SetEmbedAttr {
            index: 0,
            key: "sdtType".to_owned(),
            value: Any::from("plainText"),
        }],
        &ctx,
    )
    .unwrap();
    assert_eq!(value(&doc, "agree"), None);

    legacy(&doc, "account.reference", "REF-LEGACY");
    let reference = Position::new("body", embed_index(&doc, "account.reference"));
    doc.set_embed_attrs(
        &ctx,
        reference,
        vec![(
            "propertiesJson".to_owned(),
            Any::from(r#"{"sdtType":"checkbox"}"#),
        )],
    )
    .unwrap();
    assert_eq!(value(&doc, "account.reference"), None);
}

#[test]
fn an_inline_control_is_typed_by_its_parsed_properties_first() {
    let doc = seeded(16);
    let ctx = EditCtx::local("", "");
    doc.insert_embed(
        &ctx,
        Position::new("body", 0),
        "sdt",
        vec![
            ("tag".to_owned(), Any::from("parsed-text")),
            ("sdtType".to_owned(), Any::from("checkbox")),
            (
                "propertiesJson".to_owned(),
                Any::from(r#"{"sdtType":"plainText"}"#),
            ),
        ],
    )
    .unwrap();
    doc.insert_embed(
        &ctx,
        Position::new("body", 1),
        "sdt",
        vec![
            ("tag".to_owned(), Any::from("parsed-checkbox")),
            ("sdtType".to_owned(), Any::from("plainText")),
            (
                "propertiesJson".to_owned(),
                Any::from(r#"{"sdtType":"checkbox"}"#),
            ),
        ],
    )
    .unwrap();
    doc.insert_embed(
        &ctx,
        Position::new("body", 2),
        "sdt",
        vec![
            ("tag".to_owned(), Any::from("unparsed")),
            ("sdtType".to_owned(), Any::from("checkbox")),
            ("propertiesJson".to_owned(), Any::from("not json")),
        ],
    )
    .unwrap();
    let types: Vec<(String, String)> = listed(&doc).controls[..3]
        .iter()
        .map(|control| {
            (
                control.metadata.tag.clone().unwrap(),
                control.metadata.control_type.clone(),
            )
        })
        .collect();
    assert_eq!(
        types,
        [
            ("parsed-text", "plainText"),
            ("parsed-checkbox", "checkbox"),
            ("unparsed", "checkbox"),
        ]
        .map(|(tag, kind)| (tag.to_owned(), kind.to_owned()))
    );
    assert!(matches!(
        doc.set_embed_attrs(
            &ctx,
            Position::new("body", 0),
            vec![("value".to_owned(), checkbox_value(true))],
        ),
        Err(OpError::TextControlValue)
    ));
    for index in [1, 2] {
        doc.set_embed_attrs(
            &ctx,
            Position::new("body", index),
            vec![("value".to_owned(), checkbox_value(true))],
        )
        .unwrap();
    }
    assert!(!flagged(&doc));
    with_control(&doc, "body", "parsed-text", |map, txn| {
        map.insert(txn, "value", checkbox_value(false));
    });
    assert!(flagged(&doc));
}

#[test]
fn fills_drop_legacy_values_and_undo_keeps_them_dropped() {
    let doc = seeded(17);
    legacy(&doc, "account.reference", "REF-LEGACY");
    let undo = UndoSession::new();
    assert!(fill(&doc, &undo, "body|10000003|0", "REF-1"));
    assert_eq!(value(&doc, "account.reference"), None);
    assert_eq!(text_of(&doc, "account.reference"), "REF-1");
    assert!(undo.undo());
    assert_eq!(text_of(&doc, "account.reference"), "REF-000");
    assert_eq!(value(&doc, "account.reference"), None);
    assert!(undo.redo());
    assert_eq!(text_of(&doc, "account.reference"), "REF-1");
    assert_eq!(value(&doc, "account.reference"), None);
}

#[test]
fn equal_text_fills_drop_legacy_values_outside_history() {
    let doc = seeded(18);
    legacy(&doc, "account.reference", "REF-LEGACY");
    let undo = UndoSession::new();
    let content = |doc: &EditingDoc| {
        doc.story_segments("body")
            .unwrap()
            .into_iter()
            .find_map(|segment| match segment.content {
                SegmentContent::OtherEmbed { payload, .. }
                    if payload.get("tag") == Some(&Any::from("account.reference")) =>
                {
                    payload.get("content").cloned()
                }
                _ => None,
            })
    };
    let before = content(&doc);
    assert!(fill(&doc, &undo, "body|10000003|0", "REF-000"));
    assert_eq!(value(&doc, "account.reference"), None);
    assert_eq!(content(&doc), before);
    assert!(!flagged(&doc));
    while undo.undo() {}
    assert_eq!(value(&doc, "account.reference"), None);
    while undo.redo() {}
    assert_eq!(value(&doc, "account.reference"), None);
    assert_eq!(text_of(&doc, "account.reference"), "REF-000");
}

/// Writes the legacy state the TypeScript tests load, when `UPDATE_FIXTURES` is set: the template
/// as an older version stored it after an authored text value was set.
#[test]
fn legacy_state_fixture_loads_as_it_is() {
    if std::env::var_os("UPDATE_FIXTURES").is_some() {
        let author = seeded(18);
        legacy(&author, "account.reference", "REF-LEGACY");
        std::fs::write(LEGACY_STATE, author.encode_state_as_update_v1()).unwrap();
    }
    let state = std::fs::read(LEGACY_STATE).unwrap();
    let doc = peer(71, &state);
    assert_eq!(
        value(&doc, "account.reference"),
        Some(Any::from("REF-LEGACY"))
    );
    assert_eq!(text_of(&doc, "account.reference"), "REF-000");
}

#[test]
fn a_fill_drops_a_value_only_after_it_commits() {
    let doc = seeded(19);
    legacy(&doc, "account.reference", "REF-LEGACY");
    let origins = std::sync::Arc::new(std::sync::Mutex::new(Vec::new()));
    let seen = std::sync::Arc::clone(&origins);
    let _subscription = doc
        .yrs_doc()
        .observe_update_v1(move |txn, _| seen.lock().unwrap().push(txn.origin().cloned()))
        .unwrap();
    let undo = UndoSession::new();
    assert!(fill(&doc, &undo, "body|10000003|0", "REF-1"));
    assert_eq!(
        *origins.lock().unwrap(),
        [
            Some(yrs::Origin::from(19u64)),
            Some(yrs::Origin::from("host"))
        ]
    );
    assert_eq!(value(&doc, "account.reference"), None);
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": doc.version().as_str(),
        "steps": [{"op": "setContentControlText", "target": {"kind": "id", "controlId": "body|10000003|0"}, "text": "REF-2"}]
    }))
    .unwrap();
    legacy(&doc, "account.reference", "REF-LEGACY");
    assert!(doc.apply_edits(&request, &undo).unwrap().is_err());
    assert_eq!(
        value(&doc, "account.reference"),
        Some(Any::from("REF-LEGACY"))
    );
}
