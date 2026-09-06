use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_parse::{Bullet, ShapeNode};
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/autonumber-bullets.pptx");
const V9: &[u8] = include_bytes!("fixtures/deck-schema-v9-autonumber.update.bin");
const V11: &[u8] = include_bytes!("fixtures/deck-schema-v11-autonumber.update.bin");
const V10: &[u8] = include_bytes!("fixtures/deck-schema-v10-autonumber.update.bin");

fn restart_bullets(session: &DeckSession) -> Vec<Bullet> {
    let shape = session.package().slides[0]
        .shapes
        .iter()
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.base.id == 12 => Some(shape),
            _ => None,
        })
        .unwrap();
    shape
        .text
        .as_ref()
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.bullet.clone().unwrap())
        .collect()
}

#[test]
fn explicit_numbering_restarts_survive_text_edits_and_save() {
    let session = DeckSession::open(DECK, 30010).unwrap();
    let expected = restart_bullets(&session);
    assert!(matches!(
        expected[2],
        Bullet::AutoNumber {
            start_at: 1,
            restart: true,
            ..
        }
    ));
    let snapshot = session.snapshot().unwrap();
    let story = &snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "List 12")
        .unwrap()
        .text_stories[0];
    let restart_offset = story.paragraphs[..2]
        .iter()
        .map(|paragraph| {
            1 + paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
        })
        .sum::<u32>();
    session
        .set_paragraph_alignment(
            &EditCtx::local("numbered"),
            &story.id,
            restart_offset,
            restart_offset,
            Some("ctr"),
        )
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("numbered"),
            &story.id,
            0,
            "Edited ",
            &TextStyle::default(),
        )
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 30011).unwrap();
    assert_eq!(restart_bullets(&reopened), expected);
    let snapshot = reopened.snapshot().unwrap();
    assert!(
        snapshot.slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name == "List 12")
            .unwrap()
            .text_stories[0]
            .paragraphs[0]
            .runs[0]
            .text
            .starts_with("Edited ")
    );
}

#[test]
fn v9_and_v10_numbering_migrate_once_and_recover_explicit_restarts_from_source() {
    for (legacy, version) in [(V9, 9.0), (V10, 10.0)] {
        let doc = Doc::new();
        doc.transact_mut()
            .apply_update(Update::decode_v1(legacy).unwrap())
            .unwrap();
        let txn = doc.transact();
        assert_eq!(
            txn.get_map("pptx:meta").unwrap().get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(version)))
        );
        assert!(!String::from_utf8_lossy(legacy).contains("\"restart\""));
        assert_migrated_restarts(legacy);
    }
}

fn assert_migrated_restarts(legacy: &[u8]) {
    let migrated = DeckSession::open_from_update(legacy, 30012).unwrap();
    let update = migrated.encode_state_as_update_v1();
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let txn = doc.transact();
    assert_eq!(
        txn.get_map("pptx:meta").unwrap().get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(12.0)))
    );
    let reopened = DeckSession::open_from_update(&update, 30013).unwrap();
    assert_eq!(reopened.encode_state_as_update_v1(), update);
    assert_eq!(reopened.snapshot().unwrap(), migrated.snapshot().unwrap());
    let attached = DeckSession::open_from_update_with_source(&update, DECK, 30014).unwrap();
    let fresh = DeckSession::open(DECK, 30015).unwrap();
    assert_eq!(restart_bullets(&attached), restart_bullets(&fresh));
    assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
    let attached_update = attached.encode_state_as_update_v1();
    let recovered = DeckSession::open_from_update(&attached_update, 30016).unwrap();
    assert_eq!(restart_bullets(&recovered), restart_bullets(&fresh));
    let reattached =
        DeckSession::open_from_update_with_source(&attached_update, DECK, 30017).unwrap();
    assert_eq!(reattached.encode_state_as_update_v1(), attached_update);
    let saved = DeckSession::open(&attached.save().unwrap(), 30018).unwrap();
    assert_eq!(restart_bullets(&saved), restart_bullets(&fresh));
}

#[test]
fn main_v11_restarts_survive_spacing_migration() {
    assert!(String::from_utf8_lossy(V11).contains("\"restart\""));
    let migrated = DeckSession::open_from_update(V11, 30019).unwrap();
    let fresh = DeckSession::open(DECK, 30020).unwrap();
    assert_eq!(restart_bullets(&migrated), restart_bullets(&fresh));
    assert_eq!(migrated.snapshot().unwrap(), fresh.snapshot().unwrap());
    assert_migrated_restarts(V11);
}
