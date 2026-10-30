use std::sync::Arc;

use serde_json::Value;

use super::{EditSessionCore, EngineErrorKind, Host};
use crate::{EditTextView, ReadParagraphsRequest, StoryRange, TriState};

const BODY: &str = r#"[{"storyId":"body","paragraphs":[{"text":"First"},{"text":"Second"}]}]"#;

fn core(client_id: u64) -> EditSessionCore {
    EditSessionCore::new(client_id, Host::deterministic())
}

#[test]
fn load_json_then_load_converges() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let loaded = core(2);
    loaded.load(&source.encode_state()).unwrap();
    assert_eq!(source.encode_state_vector(), loaded.encode_state_vector());
    assert_eq!(loaded.story_ids(), vec!["body".to_owned()]);
}

#[test]
fn diff_chunks_round_trip_converges() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let target = core(2);
    let chunks = source
        .encode_diff_chunks(&target.encode_state_vector(), 64)
        .unwrap();
    assert!(!chunks.is_empty());
    target.apply_update_chunks(&chunks).unwrap();
    assert_eq!(source.encode_state_vector(), target.encode_state_vector());
}

#[test]
fn update_events_carry_origin_byte() {
    let source = core(1);
    source.load_json(BODY).unwrap();
    let state = source.encode_state();

    let mut remote = core(2);
    remote.start_update_event_observation().unwrap();
    remote.apply_update(&state).unwrap();
    let event = remote.drain_update_event();
    assert_eq!(event.first(), Some(&1));
    assert!(event.len() > 1);
    assert!(remote.drain_update_event().is_empty());

    let mut local = core(3);
    local.start_update_event_observation().unwrap();
    local.apply_local_update(&state).unwrap();
    let event = local.drain_update_event();
    assert_eq!(event.first(), Some(&0));
    assert!(event.len() > 1);
}

#[test]
fn malformed_update_error_matches_shim_text() {
    let session = core(1);
    let error = session.apply_update(&[255, 255, 255]).unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert!(
        error.message.starts_with("invalid yrs update: "),
        "{}",
        error.message
    );
}

#[test]
fn version_nonce_depends_on_host_entropy() {
    let one = EditSessionCore::new(7, Host::with_entropy(Arc::new(|| 1)));
    let two = EditSessionCore::new(7, Host::with_entropy(Arc::new(|| 2)));
    assert_ne!(one.engine.doc().version(), two.engine.doc().version());
}

fn paragraph_texts(session: &EditSessionCore) -> Vec<String> {
    session
        .engine
        .doc()
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

#[test]
fn text_edits_marks_and_raw_ops_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let inserted: Value = serde_json::from_str(
        &session
            .insert_text("body", para_id, 5, "!", None, None)
            .unwrap(),
    )
    .unwrap();
    assert!(inserted["revisionId"].is_null());
    assert_eq!(inserted["range"]["story"], "body");
    assert_eq!(paragraph_texts(&session), ["First!", "Second"]);
    let deleted: Value = serde_json::from_str(
        &session
            .delete_range("body", para_id, 5, para_id, 6, None, None)
            .unwrap(),
    )
    .unwrap();
    assert!(deleted["revisionId"].is_null());
    assert_eq!(deleted["range"]["story"], "body");
    assert_eq!(paragraph_texts(&session), ["First", "Second"]);

    session
        .toggle_mark("body", para_id, 0, para_id, 5, r#"{"type":"bold"}"#)
        .unwrap();
    assert_eq!(
        session
            .engine
            .doc()
            .selection_context(&StoryRange::new("body", 0, 5))
            .unwrap()
            .bold,
        TriState::On
    );
    session
        .apply_raw_ops("body", r#"[{"op":"insert","index":0,"text":"+"}]"#)
        .unwrap();
    assert_eq!(paragraph_texts(&session), ["+First", "Second"]);
}

#[test]
fn table_edits_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let inserted: Value = serde_json::from_str(
        &session
            .insert_table("body", para_id, 5, 1, 2, None, None)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(inserted["rows"], 1);
    assert_eq!(inserted["columns"], 2);
    let row: Value = serde_json::from_str(
        &session
            .insert_row(
                r#"{"story":"body","tableIndex":0,"row":0,"column":0}"#,
                true,
                None,
                None,
            )
            .unwrap(),
    )
    .unwrap();
    assert_eq!(row["rows"], 2);
    assert_eq!(row["columns"], 2);
    let deleted: Value = serde_json::from_str(
        &session
            .delete_table(&inserted["table"].to_string())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(deleted["deletedTable"], true);
    assert_eq!(session.story_ids(), ["body"]);
}

#[test]
fn incomplete_suggesting_author_matches_shim_error() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let before = session.encode_state();
    let error = session
        .insert_text("body", para_id, 0, "!", Some("author".to_owned()), None)
        .unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(
        error.message,
        "suggesting requires both an author name and an ISO date"
    );
    assert_eq!(session.encode_state(), before);
}
