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

#[test]
fn selection_and_undo_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    session
        .set_selection("body", para_id, 1, para_id, 4)
        .unwrap();
    let selection: Value = serde_json::from_str(&session.selection().unwrap()).unwrap();
    assert_eq!(
        selection,
        serde_json::json!({
            "anchor": { "story": "body", "paraId": para_id, "offset": 1 },
            "head": { "story": "body", "paraId": para_id, "offset": 4 },
        })
    );

    session.track_undo();
    session
        .insert_text("body", para_id, 5, "!", None, None)
        .unwrap();
    assert_eq!(paragraph_texts(&session), ["First!", "Second"]);
    assert!(session.undo());
    assert_eq!(paragraph_texts(&session), ["First", "Second"]);
    assert!(session.can_redo());
}

#[test]
fn comments_reads_and_identities_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let ranges = serde_json::json!([{
        "story": "body", "startPara": para_id, "startOffset": 0,
        "endPara": para_id, "endOffset": 5,
    }]);
    let added: Value = serde_json::from_str(
        &session
            .add_comment(&ranges.to_string(), "Reviewer", "2026-01-01", r#""Note""#)
            .unwrap(),
    )
    .unwrap();
    let comments: Value = serde_json::from_str(&session.list_comments().unwrap()).unwrap();
    assert_eq!(comments.as_array().unwrap().len(), 1);
    assert_eq!(comments[0]["id"], added["commentId"]);
    assert_eq!(comments[0]["body"], "Note");

    let read: Value = serde_json::from_str(
        &session
            .read_paragraphs_json(r#"{"story":"body","view":"accepted"}"#)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(read["ok"], true);
    let paragraphs = read["paragraphs"].as_array().unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0]["text"], "First");
    assert_eq!(paragraphs[1]["text"], "Second");
    let identities: Value = serde_json::from_str(&session.paragraph_identities().unwrap()).unwrap();
    let identities = identities["paragraphs"].as_array().unwrap();
    assert_eq!(identities.len(), paragraphs.len());
    for (identity, paragraph) in identities.iter().zip(paragraphs) {
        assert_eq!(identity["session"]["paraId"], paragraph["paraId"]);
    }
}

#[test]
fn unknown_selection_story_matches_shim_error() {
    let error = core(1)
        .set_selection("missing", "p1", 0, "p1", 0)
        .unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(error.message, "story \"missing\" was not found");
}

#[test]
fn fonts_layout_frames_and_resident_input_run_without_wasm() {
    let session = core(1);
    let seeded: Value = serde_json::from_str(&session.load_json(BODY).unwrap()).unwrap();
    let para_id = seeded["body"][0].as_str().unwrap();
    let ordinary = core(2);
    ordinary.load(&session.encode_state()).unwrap();
    let font = session
        .register_measure_font(include_bytes!(
            "../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "properties": {
                "pageWidth": 4320,
                "pageHeight": 2880,
                "marginTop": 300,
                "marginRight": 300,
                "marginBottom": 300,
                "marginLeft": 300
            }
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let layout: Value = serde_json::from_str(
        &session
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    assert!(!layout["layout"]["pages"].as_array().unwrap().is_empty());
    let frame = session
        .build_display_list_frame(
            &serde_json::json!({"fontChains": {"calibri|0|0": [font]}}).to_string(),
            0.0,
        )
        .unwrap();
    assert!(!frame.is_empty());
    session
        .set_selection("body", para_id, 2, para_id, 2)
        .unwrap();
    let caret: Value =
        serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
    let rect = &caret["caretRect"];
    let hit: Value = serde_json::from_str(
        &session
            .display_hit_test_regions_json(
                rect["pageIndex"].as_u64().unwrap() as u32,
                rect["x"].as_f64().unwrap(),
                rect["y"].as_f64().unwrap() + rect["height"].as_f64().unwrap() / 2.0,
            )
            .unwrap(),
    )
    .unwrap();
    assert_eq!(hit["region"], "body");
    assert_eq!(hit["target"], "text");
    let hit_pos = hit["pos"].as_i64().unwrap();
    let hit_caret = session
        .engine
        .with_display_list(|list| docx_layout::hit::caret_rect(list, hit_pos))
        .unwrap()
        .unwrap();
    let hit_caret = serde_json::to_value(hit_caret).unwrap();
    for key in ["pageIndex", "x", "y", "height"] {
        assert_eq!(hit_caret[key], rect[key], "{key}");
    }

    let receipt: Value = serde_json::from_str(
        &ordinary
            .insert_text("body", para_id, 2, "!", None, None)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        receipt,
        serde_json::json!({
            "revisionId": null,
            "range": {
                "story": "body",
                "start": {"paraId": para_id, "offset": 2},
                "end": {"paraId": para_id, "offset": 3}
            }
        })
    );
    let frame = session
        .apply_input("!", caret["frameEpoch"].as_f64().unwrap())
        .unwrap();
    assert!(!frame.is_empty());
    assert_eq!(paragraph_texts(&session), ["Fi!rst", "Second"]);
    assert_eq!(paragraph_texts(&session), paragraph_texts(&ordinary));
    let selection: Value = serde_json::from_str(&session.selection().unwrap()).unwrap();
    assert_eq!(selection["head"]["story"], receipt["range"]["story"]);
    assert_eq!(
        selection["head"]["paraId"],
        receipt["range"]["end"]["paraId"]
    );
    assert_eq!(
        selection["head"]["offset"],
        receipt["range"]["end"]["offset"]
    );
    assert_eq!(selection["anchor"], selection["head"]);
}

#[test]
fn invalid_resident_frame_epoch_matches_shim_error_before_mutation() {
    let session = core(1);
    session.load_json(BODY).unwrap();
    let before = session.encode_state();
    let error = session.apply_input("!", -1.0).unwrap_err();
    assert_eq!(error.kind, EngineErrorKind::Message);
    assert_eq!(
        error.message,
        "expected_frame_epoch must be a non-negative safe integer"
    );
    assert_eq!(session.encode_state(), before);
}
