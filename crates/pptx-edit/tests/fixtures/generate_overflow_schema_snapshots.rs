use std::{env, fs, path::Path};

use pptx_edit::DeckSession;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update};

fn main() {
    let args = env::args().collect::<Vec<_>>();
    let root = Path::new(&args[1]);
    let version: u32 = args[2].parse().unwrap();
    assert!(matches!(version, 14 | 15));
    for (source, name) in [
        (
            "crates/pptx-edit/tests/fixtures/chart-text-overflow.pptx",
            "chart-text-overflow",
        ),
        (
            "crates/pptx-render/tests/fixtures/text-overflow.pptx",
            "text-overflow",
        ),
    ] {
        if version == 14 && name == "text-overflow" {
            continue;
        }
        let session = DeckSession::open(&fs::read(root.join(source)).unwrap(), 29501).unwrap();
        let update = session.encode_state_as_update_v1();
        let doc = Doc::new();
        doc.transact_mut()
            .apply_update(Update::decode_v1(&update).unwrap())
            .unwrap();
        let txn = doc.transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(version as f64)))
        );
        let Some(Out::Any(Any::Buffer(json))) = meta.get(&txn, "packageJson") else {
            panic!("missing package JSON");
        };
        let json = std::str::from_utf8(&json).unwrap();
        assert!(!json.contains("verticalOverflow"));
        assert!(!json.contains("horizontalOverflow"));
        if name == "chart-text-overflow" {
            assert_eq!(json.contains("\"kind\":\"pattern\""), version == 15);
            assert_eq!(json.contains("\"line\":{\"none\":true}"), version == 15);
        }
        let reopened = DeckSession::open_from_update(&update, 29502).unwrap();
        assert_eq!(reopened.encode_state_as_update_v1(), update);
        fs::write(
            root.join(format!(
                "crates/pptx-edit/tests/fixtures/deck-schema-v{version}-{name}.update.bin"
            )),
            update,
        )
        .unwrap();
    }
}
