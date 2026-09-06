use std::{env, fs, path::Path};

use pptx_edit::DeckSession;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update};

fn main() {
    let args = env::args().collect::<Vec<_>>();
    let root = Path::new(&args[1]);
    let source = fs::read(root.join("crates/pptx-render/tests/fixtures/blip-effects.pptx")).unwrap();
    let session = DeckSession::open(&source, 312).unwrap();
    let update = session.encode_state_as_update_v1();
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let txn = doc.transact();
    let meta = txn.get_map("pptx:meta").unwrap();
    assert_eq!(
        meta.get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(16.0)))
    );
    let Some(Out::Any(Any::Buffer(json))) = meta.get(&txn, "packageJson") else {
        panic!("missing package JSON");
    };
    assert!(!std::str::from_utf8(&json).unwrap().contains("\"effects\""));
    let shapes = txn.get_map("pptx:shapes").unwrap();
    assert!(shapes.iter(&txn).all(|(_, shape)| {
        !shape
            .cast::<yrs::MapRef>()
            .unwrap()
            .contains_key(&txn, "blipEffectsJson")
    }));
    let reopened = DeckSession::open_from_update(&update, 313).unwrap();
    assert_eq!(reopened.encode_state_as_update_v1(), update);
    fs::write(
        root.join("crates/pptx-edit/tests/fixtures/blip-effects-main-v16.update.bin"),
        update,
    )
    .unwrap();
}
