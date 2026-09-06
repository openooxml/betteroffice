use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};
use yrs::{Any, Map, Out, ReadTxn, Transact};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/run-spacing.pptx");
const V10: &[u8] = include_bytes!("fixtures/run-spacing-main-v10.update.bin");
const STORY: &str = "story:slide:0:256:shape:0:0";

#[test]
fn v10_tracking_recovers_with_source_and_preserves_edits() {
    let migrated = DeckSession::open_from_update_with_source(V10, DECK, 32501).unwrap();
    let fresh = DeckSession::open(DECK, 32502).unwrap();
    assert_eq!(migrated.snapshot().unwrap(), fresh.snapshot().unwrap());
    assert_eq!(
        ooxml_opc::unzip_parts(&migrated.save().unwrap()).unwrap(),
        ooxml_opc::unzip_parts(DECK).unwrap()
    );
    let detached = DeckSession::open_from_update(V10, 32503).unwrap();
    for session in [&detached, &fresh] {
        session
            .insert_text(
                &EditCtx::local("test"),
                STORY,
                0,
                "😀 ",
                &TextStyle::default(),
            )
            .unwrap();
        session
            .format_text(
                &EditCtx::local("test"),
                STORY,
                5,
                6,
                &TextStylePatch {
                    spacing_pt: Some(0.0),
                    ..Default::default()
                },
            )
            .unwrap();
    }
    let recovered = DeckSession::open_from_update_with_source(
        &detached.encode_state_as_update_v1(),
        DECK,
        32504,
    )
    .unwrap();
    assert_eq!(recovered.snapshot().unwrap(), fresh.snapshot().unwrap());
    let update = recovered.encode_state_as_update_v1();
    let reopened = DeckSession::open_from_update(&update, 32505).unwrap();
    assert_eq!(reopened.encode_state_as_update_v1(), update);
    assert_eq!(reopened.snapshot().unwrap(), recovered.snapshot().unwrap());
    let txn = reopened.yrs_doc().transact();
    let meta = txn.get_map("pptx:meta").unwrap();
    assert_eq!(
        meta.get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(11.0)))
    );
    assert!(meta.get(&txn, "spacingPendingSource").is_none());
}

#[test]
fn tracking_formatting_round_trips_and_rejects_invalid_values() {
    let session = DeckSession::open(DECK, 32506).unwrap();
    let context = EditCtx::local("test");
    let end = session.story(STORY).unwrap().length - 1;
    for spacing in [3.0, -1.0, 0.0] {
        session
            .format_text(
                &context,
                STORY,
                0,
                end,
                &TextStylePatch {
                    spacing_pt: Some(spacing),
                    ..Default::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 32507).unwrap();
        assert_eq!(
            reopened.story(STORY).unwrap(),
            session.story(STORY).unwrap()
        );
        assert_eq!(
            reopened.story(STORY).unwrap().paragraphs[0].runs[0]
                .style
                .spacing_pt,
            Some(spacing)
        );
    }
    let before = session.encode_state_as_update_v1();
    for spacing in [
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
        -4000.01,
        4000.01,
    ] {
        assert!(
            session
                .format_text(
                    &context,
                    STORY,
                    0,
                    1,
                    &TextStylePatch {
                        spacing_pt: Some(spacing),
                        ..Default::default()
                    }
                )
                .is_err()
        );
        assert!(
            session
                .insert_text(
                    &context,
                    STORY,
                    0,
                    "X",
                    &TextStyle {
                        spacing_pt: Some(spacing),
                        ..Default::default()
                    }
                )
                .is_err()
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}
