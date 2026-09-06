use std::{env, fs, path::PathBuf};

use pptx_edit::DeckSession;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update, updates::decoder::Decode};

fn main() {
    let root = PathBuf::from(env::args().nth(1).unwrap());
    let bytes = fs::read(root.join("crates/pptx-parse/tests/fixtures/gradient-outline.pptx")).unwrap();
    let session = DeckSession::open(&bytes, 322).unwrap();
    let update = session.encode_state_as_update_v1();
    let doc = Doc::with_client_id(323);
    doc.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let txn = doc.transact();
    let meta = txn.get_map("pptx:meta").unwrap();
    assert_eq!(
        meta.get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(13.0)))
    );
    assert!(meta.get(&txn, "outlineGradientsPendingSource").is_none());
    let reopened = DeckSession::open_from_update(&update, 324).unwrap();
    assert_eq!(reopened.encode_state_as_update_v1(), update);
    fs::write(
        root.join("crates/pptx-edit/tests/fixtures/gradient-outline-main-v13.update.bin"),
        update,
    )
    .unwrap();
}
