//! `deck-schema-v1.update.bin` was produced by release 4bdccdd: it opens
//! `betteroffice-demo.pptx`, adds a text box and edits its story, then persists
//! `encode_state_as_update_v1()`.

use pptx_edit::{DeckSession, DeckSnapshot, EditCtx, EditError, TextStyle};
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, MapRef, Out, ReadTxn, StateVector, Transact, Update};

const V1_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v1.update.bin");
const META: &str = "pptx:meta";
const SHAPE_ID: &str = "shape:4242:0";
const STORY_ID: &str = "story:shape:4242:0:0";

#[test]
fn released_v1_snapshot_migrates_and_round_trips_as_v2() {
    assert_eq!(stamped_version(V1_UPDATE), Some(1.0));

    let session = DeckSession::open_from_update(V1_UPDATE, 901).unwrap();
    assert_v1_content(&session);

    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.0));
    assert!(
        package_json(&migrated).contains("\"charts\""),
        "the migrated package must carry the v2 chart field"
    );

    let reopened = DeckSession::open_from_update(&migrated, 902).unwrap();
    assert_v1_content(&reopened);
    assert_eq!(
        snapshot_shape_ids(&session.snapshot().unwrap()),
        snapshot_shape_ids(&reopened.snapshot().unwrap())
    );
    assert_eq!(
        reopened.encode_state_as_update_v1().len(),
        migrated.len(),
        "reopening a v2 snapshot must not migrate again"
    );
}

#[test]
fn a_migrated_session_still_edits() {
    let session = DeckSession::open_from_update(V1_UPDATE, 903).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            STORY_ID,
            0,
            "re-",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
    let reopened =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 904).unwrap();
    assert_eq!(
        reopened.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
}

#[test]
fn two_clients_migrating_the_same_v1_snapshot_converge() {
    let left = DeckSession::open_from_update(V1_UPDATE, 907).unwrap();
    let right = DeckSession::open_from_update(V1_UPDATE, 908).unwrap();

    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(
        stamped_version(&left.encode_state_as_update_v1()),
        Some(2.0)
    );
    assert_eq!(
        package_json(&left.encode_state_as_update_v1()),
        package_json(&right.encode_state_as_update_v1())
    );
}

#[test]
fn unmigratable_schema_versions_stay_rejected() {
    for version in [0.0, 1.5, 3.0] {
        assert!(
            matches!(
                DeckSession::open_from_update(&restamped(V1_UPDATE, Some(version)), 905),
                Err(EditError::InvalidState(message))
                    if message == "unsupported deck schema version"
            ),
            "schema version {version} must be rejected"
        );
    }
    assert!(matches!(
        DeckSession::open_from_update(&restamped(V1_UPDATE, None), 906),
        Err(EditError::InvalidState(message))
            if message == "unsupported deck schema version"
    ));
}

fn assert_v1_content(session: &DeckSession) {
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.width_emu, 12_192_000);
    assert_eq!(snapshot.height_emu, 6_858_000);
    assert_eq!(snapshot.slides.len(), 3);
    assert_eq!(snapshot.slides[0].id, "slide:0:256");
    assert!(
        snapshot_shape_ids(&snapshot)
            .iter()
            .any(|id| id == SHAPE_ID)
    );
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "edited persisted on v1"
    );
    assert!(session.package().charts.is_empty());
}

fn snapshot_shape_ids(snapshot: &DeckSnapshot) -> Vec<String> {
    snapshot
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .map(|shape| shape.id.clone())
        .collect()
}

fn hydrated(update: &[u8]) -> Doc {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    doc
}

fn meta(doc: &Doc) -> MapRef {
    doc.transact().get_map(META).unwrap()
}

fn stamped_version(update: &[u8]) -> Option<f64> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "schemaVersion") {
        Some(Out::Any(Any::Number(value))) => Some(value),
        _ => None,
    }
}

fn package_json(update: &[u8]) -> String {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "packageJson") {
        Some(Out::Any(Any::Buffer(bytes))) => String::from_utf8(bytes.to_vec()).unwrap(),
        _ => panic!("missing packageJson"),
    }
}

fn restamped(update: &[u8], version: Option<f64>) -> Vec<u8> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    {
        let mut txn = doc.transact_mut();
        match version {
            Some(version) => {
                meta.insert(&mut txn, "schemaVersion", version);
            }
            None => {
                meta.remove(&mut txn, "schemaVersion");
            }
        }
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}
