use std::{env, fs, path::PathBuf};

use pptx_edit::DeckSession;
use yrs::{Any, Map, Out, ReadTxn, Transact};

fn main() {
    let root = PathBuf::from(env::args().nth(1).unwrap());
    for (source, name) in [
        (
            "crates/pptx-render/tests/fixtures/outer-shadow.pptx",
            "outer-shadow",
        ),
        (
            "crates/pptx-render/tests/fixtures/outer-shadow-scale.pptx",
            "outer-shadow-scale",
        ),
        (
            "crates/pptx-edit/tests/fixtures/blip-shadow.pptx",
            "blip-shadow",
        ),
    ] {
        let session = DeckSession::open(&fs::read(root.join(source)).unwrap(), 33400).unwrap();
        let txn = session.yrs_doc().transact();
        assert_eq!(
            txn.get_map("pptx:meta").unwrap().get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(17.0)))
        );
        fs::write(
            root.join(format!(
                "crates/pptx-edit/tests/fixtures/{name}-main-v17.update.bin"
            )),
            session.encode_state_as_update_v1(),
        )
        .unwrap();
    }
}
