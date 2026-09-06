use std::{env, fs, path::PathBuf};

use pptx_edit::DeckSession;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update, updates::decoder::Decode};

fn main() {
    let root = PathBuf::from(env::args().nth(1).unwrap());
    for (source, output) in [
        ("autonumber-bullets", "deck-schema-v10-autonumber"),
        ("text-baseline-script", "deck-schema-v10-baseline"),
    ] {
        let bytes = fs::read(root.join(format!("crates/pptx-render/tests/fixtures/{source}.pptx")))
            .unwrap();
        let session = DeckSession::open(&bytes, 30001).unwrap();
        let update = session.encode_state_as_update_v1();
        let doc = Doc::new();
        doc.transact_mut()
            .apply_update(Update::decode_v1(&update).unwrap())
            .unwrap();
        let txn = doc.transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(10.0)))
        );
        assert!(meta.get(&txn, "baselinesPendingSource").is_none());
        let reopened = DeckSession::open_from_update(&update, 30002).unwrap();
        assert_eq!(reopened.encode_state_as_update_v1(), update);
        fs::write(
            root.join(format!(
                "crates/pptx-edit/tests/fixtures/{output}.update.bin"
            )),
            update,
        )
        .unwrap();
    }
}
