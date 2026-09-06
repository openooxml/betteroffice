use pptx_edit::{DeckSession, EditCtx};
use yrs::{Any, Map, Out, ReadTxn, Transact};

const SOURCE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/blip-effects.pptx");
const MAIN_V9: &[u8] = include_bytes!("fixtures/blip-effects-main-v9.update.bin");

const MAIN_V10: &[u8] = include_bytes!("fixtures/blip-effects-main-v10.update.bin");

#[test]
fn main_snapshots_import_blip_effects_without_losing_edits_or_source_parts() {
    for update in [MAIN_V9, MAIN_V10] {
        let migrated = DeckSession::open_from_update(update, 31201).unwrap();
        let txn = migrated.yrs_doc().transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(16.0)))
        );
        drop(txn);
        let initial = migrated.snapshot().unwrap();
        assert!(
            initial.slides[0]
                .shapes
                .iter()
                .all(|shape| shape.blip_effects.is_empty())
        );
        let attached = DeckSession::open_from_update_with_source(update, SOURCE, 31202).unwrap();
        let fresh = DeckSession::open(SOURCE, 31203).unwrap();
        assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(attached.package(), fresh.package());
        assert_eq!(
            attached.snapshot().unwrap().slides[0].shapes[1]
                .blip_effects
                .len(),
            1
        );
        assert_eq!(
            ooxml_opc::unzip_parts(&attached.save().unwrap()).unwrap(),
            ooxml_opc::unzip_parts(SOURCE).unwrap()
        );
        let update = attached.encode_state_as_update_v1();
        let reopened = DeckSession::open_from_update(&update, 31204).unwrap();
        assert_eq!(reopened.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(reopened.encode_state_as_update_v1(), update);
        assert_eq!(
            serde_json::to_vec(reopened.package()).unwrap(),
            serde_json::to_vec(fresh.package()).unwrap()
        );
        let reattached = DeckSession::open_from_update_with_source(&update, SOURCE, 31205).unwrap();
        assert_eq!(reattached.encode_state_as_update_v1(), update);
        let target = &initial.slides[0].shapes[1];
        migrated
            .move_shape(
                &EditCtx::local("test"),
                &initial.slides[0].id,
                &target.id,
                900,
                1000,
            )
            .unwrap();
        let edited = DeckSession::open_from_update_with_source(
            &migrated.encode_state_as_update_v1(),
            SOURCE,
            31206,
        )
        .unwrap();
        let snapshot = edited.snapshot().unwrap();
        assert_eq!(
            (
                snapshot.slides[0].shapes[1].x,
                snapshot.slides[0].shapes[1].y
            ),
            (900, 1000)
        );
        assert_eq!(
            snapshot.slides[0].shapes[1].blip_effects,
            fresh.snapshot().unwrap().slides[0].shapes[1].blip_effects
        );
        let saved = edited.save().unwrap();
        let reopened = DeckSession::open(&saved, 31207).unwrap();
        assert_eq!(
            reopened.snapshot().unwrap().slides[0].shapes[1].blip_effects,
            snapshot.slides[0].shapes[1].blip_effects
        );
        assert!(!edited.can_undo());
    }
}
