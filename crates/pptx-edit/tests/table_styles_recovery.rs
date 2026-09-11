use pptx_edit::DeckSession;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

/// The stored package as a document written before table styles were parsed:
/// the same fingerprint, with the field absent from `packageJson`.
fn stored_without_table_styles() -> Vec<u8> {
    let update = DeckSession::open(DECK, 5201)
        .unwrap()
        .encode_state_as_update_v1();
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let mut txn = doc.transact_mut();
    let meta = txn.get_map("pptx:meta").unwrap();
    let Some(Out::Any(Any::Buffer(bytes))) = meta.get(&txn, "packageJson") else {
        panic!("expected a packageJson buffer");
    };
    let mut package: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
    assert!(package.get("tableStyles").is_some());
    package.as_object_mut().unwrap().remove("tableStyles");
    let stripped = serde_json::to_vec(&package).unwrap();
    meta.insert(
        &mut txn,
        "packageJson",
        Any::Buffer(std::sync::Arc::from(stripped)),
    );
    drop(txn);
    doc.transact()
        .encode_state_as_update_v1(&Default::default())
}

#[test]
fn a_document_stored_without_table_styles_recovers_them_from_its_source() {
    let stored = stored_without_table_styles();

    let detached = DeckSession::open_from_update(&stored, 5203).unwrap();
    assert!(
        detached.package().table_styles.styles.is_empty(),
        "the stored document must start without table styles"
    );

    // Reattaching replaces the in-memory package outright, so the question is
    // what it persisted: reopen the result with no source at all.
    let reattached = DeckSession::open_from_update_with_source(&stored, DECK, 5204).unwrap();
    let persisted = reattached.encode_state_as_update_v1();

    let reopened = DeckSession::open_from_update(&persisted, 5205).unwrap();
    let recovered = &reopened.package().table_styles;
    assert!(
        !recovered.styles.is_empty(),
        "a source reattachment must persist the styles the stored package never carried"
    );
    assert!(recovered.default_style_id.is_some());
}
