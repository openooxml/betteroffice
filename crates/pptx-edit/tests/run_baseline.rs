use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update, updates::decoder::Decode};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/text-baseline-script.pptx");
const V9: &[u8] = include_bytes!("fixtures/run-baseline-main-v9.update.bin");
const EDITED_V9: &[u8] = include_bytes!("fixtures/run-baseline-edited-main-v9.update.bin");
const STORY: &str = "story:slide:0:256:shape:3:0";

#[test]
fn v10_and_current_main_v11_baselines_survive_numbering_and_spacing_migrations() {
    for update in [
        include_bytes!("fixtures/deck-schema-v10-baseline.update.bin").as_slice(),
        include_bytes!("fixtures/deck-schema-v11-baseline.update.bin").as_slice(),
    ] {
        let migrated = DeckSession::open_from_update(update, 33110).unwrap();
        let fresh = DeckSession::open(DECK, 33111).unwrap();
        assert_eq!(migrated.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(
            serde_json::to_value(migrated.package()).unwrap(),
            serde_json::to_value(fresh.package()).unwrap()
        );
        let current = migrated.encode_state_as_update_v1();
        let reattached = DeckSession::open_from_update_with_source(&current, DECK, 33112).unwrap();
        assert_eq!(reattached.encode_state_as_update_v1(), current);
        assert_eq!(
            ooxml_opc::unzip_parts(&reattached.save().unwrap()).unwrap(),
            ooxml_opc::unzip_parts(DECK).unwrap()
        );
    }
}

#[test]
fn v9_baselines_recover_with_source_and_preserve_edits() {
    for (update, prefix) in [(V9, ""), (EDITED_V9, "😀 ")] {
        let session = DeckSession::open_from_update_with_source(update, DECK, 33101).unwrap();
        let fresh = DeckSession::open(DECK, 33102).unwrap();
        if !prefix.is_empty() {
            fresh
                .insert_text(
                    &EditCtx::local("test"),
                    STORY,
                    0,
                    prefix,
                    &TextStyle::default(),
                )
                .unwrap();
        }
        assert_eq!(session.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(
            serde_json::to_value(session.package()).unwrap(),
            serde_json::to_value(fresh.package()).unwrap()
        );
        let detached = DeckSession::open_from_update(update, 33107).unwrap();
        for candidate in [&detached, &fresh] {
            let context = EditCtx::local("test");
            candidate
                .insert_text(&context, STORY, 0, "X", &TextStyle::default())
                .unwrap();
            let end = candidate.story(STORY).unwrap().length - 1;
            candidate
                .insert_text(&context, STORY, end, "Y", &TextStyle::default())
                .unwrap();
            let raised = prefix.encode_utf16().count() as u32 + 7;
            candidate
                .format_text(
                    &context,
                    STORY,
                    raised,
                    raised + 1,
                    &TextStylePatch {
                        baseline_pct: Some(0.0),
                        ..Default::default()
                    },
                )
                .unwrap();
        }
        let recovered = DeckSession::open_from_update_with_source(
            &detached.encode_state_as_update_v1(),
            DECK,
            33108,
        )
        .unwrap();
        assert_eq!(recovered.snapshot().unwrap(), fresh.snapshot().unwrap());
        let fresh = DeckSession::open(DECK, 33109).unwrap();
        if !prefix.is_empty() {
            fresh
                .insert_text(
                    &EditCtx::local("test"),
                    STORY,
                    0,
                    prefix,
                    &TextStyle::default(),
                )
                .unwrap();
        }
        let current = session.encode_state_as_update_v1();
        let doc = Doc::new();
        doc.transact_mut()
            .apply_update(Update::decode_v1(&current).unwrap())
            .unwrap();
        let txn = doc.transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(12.0)))
        );
        assert!(meta.get(&txn, "baselinesPendingSource").is_none());
        let reopened = DeckSession::open_from_update(&current, 33103).unwrap();
        assert_eq!(reopened.encode_state_as_update_v1(), current);
        assert_eq!(reopened.snapshot().unwrap(), session.snapshot().unwrap());
        assert_eq!(
            serde_json::to_value(reopened.package()).unwrap(),
            serde_json::to_value(fresh.package()).unwrap()
        );
        let reattached = DeckSession::open_from_update_with_source(&current, DECK, 33104).unwrap();
        assert_eq!(reattached.encode_state_as_update_v1(), current);
        assert_eq!(
            ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap(),
            ooxml_opc::unzip_parts(&fresh.save().unwrap()).unwrap()
        );
        if prefix.is_empty() {
            assert_eq!(
                ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap(),
                ooxml_opc::unzip_parts(DECK).unwrap()
            );
        }
    }
}

#[test]
fn baseline_formatting_round_trips_and_rejects_invalid_values() {
    let session = DeckSession::open(DECK, 33105).unwrap();
    let context = EditCtx::local("test");
    for baseline in [150.0, -150.0, 0.0] {
        session
            .format_text(
                &context,
                STORY,
                6,
                7,
                &TextStylePatch {
                    baseline_pct: Some(baseline),
                    ..Default::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 33106).unwrap();
        assert_eq!(
            reopened.story(STORY).unwrap(),
            session.story(STORY).unwrap()
        );
        assert_eq!(
            reopened.story(STORY).unwrap().paragraphs[0].runs[1]
                .style
                .baseline_pct,
            Some(baseline)
        );
    }
    let before = session.encode_state_as_update_v1();
    for baseline in [
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
        2147483.648,
        -2147483.649,
    ] {
        assert!(
            session
                .format_text(
                    &context,
                    STORY,
                    6,
                    7,
                    &TextStylePatch {
                        baseline_pct: Some(baseline),
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
                        baseline_pct: Some(baseline),
                        ..Default::default()
                    }
                )
                .is_err()
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn absent_baseline_preserves_legacy_style_json() {
    let json = r#"{"bold":null,"italic":null,"fontSizePt":null,"color":null,"fontFamily":null,"underline":null}"#;
    assert_eq!(
        serde_json::to_string(&serde_json::from_str::<TextStyle>(json).unwrap()).unwrap(),
        json
    );
    assert_eq!(
        serde_json::to_string(&serde_json::from_str::<TextStylePatch>(json).unwrap()).unwrap(),
        json
    );
}
