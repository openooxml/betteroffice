use std::{env, fs, path::PathBuf};

use pptx_edit::{DeckSession, EditCtx, TextStyle};
use yrs::{Any, Doc, Map, Out, ReadTxn, StateVector, Transact, Update, updates::decoder::Decode};

fn main() {
    let root = PathBuf::from(env::args().nth(1).unwrap());
    for (crate_name, source, name) in [
        ("pptx-render", "line-spacing", "line-spacing"),
        ("pptx-render", "autonumber-bullets", "autonumber"),
        ("pptx-render", "text-baseline-script", "baseline"),
        ("pptx-parse", "picture-fill", "picture-fill"),
        ("pptx-render", "chart-space-fill", "chart-space-fill"),
    ] {
        let bytes =
            fs::read(root.join(format!("crates/{crate_name}/tests/fixtures/{source}.pptx")))
                .unwrap();
        let session = DeckSession::open(&bytes, 31401).unwrap();
        if name == "line-spacing" {
            let story = &session.snapshot().unwrap().slides[0].shapes[0].text_stories[0];
            session
                .insert_text(
                    &EditCtx::local("fixture"),
                    &story.id,
                    0,
                    "Edited ",
                    &TextStyle::default(),
                )
                .unwrap();
        }
        let update = session.encode_state_as_update_v1();
        let doc = Doc::with_client_id(31402);
        doc.transact_mut()
            .apply_update(Update::decode_v1(&update).unwrap())
            .unwrap();
        let txn = doc.transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(13.0)))
        );
        assert!(meta.get(&txn, "baselinesPendingSource").is_none());
        let reopened = DeckSession::open_from_update(&update, 31403).unwrap();
        assert_eq!(reopened.encode_state_as_update_v1(), update);
        fs::write(
            root.join(format!(
                "crates/pptx-edit/tests/fixtures/deck-schema-v13-{name}.update.bin"
            )),
            update,
        )
        .unwrap();
        drop(txn);
        if matches!(name, "autonumber" | "baseline") {
            let mut package = serde_json::to_value(session.package()).unwrap();
            remove_post_v10_properties(&mut package);
            let package = serde_json::from_value(package).unwrap();
            let legacy = DeckSession::from_package_with_source(package, &bytes, 31404).unwrap();
            let doc = Doc::with_client_id(31405);
            let mut txn = doc.transact_mut();
            txn.apply_update(Update::decode_v1(&legacy.encode_state_as_update_v1()).unwrap())
                .unwrap();
            let meta = txn.get_map("pptx:meta").unwrap();
            meta.insert(&mut txn, "schemaVersion", 10.0);
            let update = txn.encode_state_as_update_v1(&StateVector::default());
            let migrated = DeckSession::open_from_update(&update, 31406).unwrap();
            assert_eq!(migrated.snapshot().unwrap(), legacy.snapshot().unwrap());
            fs::write(
                root.join(format!(
                    "crates/pptx-edit/tests/fixtures/deck-schema-v10-{name}.update.bin"
                )),
                update,
            )
            .unwrap();
        }
    }
}

fn remove_post_v10_properties(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(object) => {
            object.remove("restart");
            object.remove("lineSpacing");
            object.remove("compatLineSpacing");
            object.remove("pictureFill");
            for value in object.values_mut() {
                remove_post_v10_properties(value);
            }
        }
        serde_json::Value::Array(array) => {
            for value in array {
                remove_post_v10_properties(value);
            }
        }
        _ => {}
    }
}
