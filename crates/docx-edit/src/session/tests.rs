use std::sync::Arc;

use super::{EditSessionCore, EngineErrorKind, Host};

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
