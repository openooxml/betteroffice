use std::{env, fs, path::Path};

use pptx_edit::{DeckSession, EditCtx};
use yrs::{Map, ReadTxn, Transact};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let branch = Path::new(&args[1]);
    let fixtures = branch.join("crates/pptx-edit/tests/fixtures");
    for name in ["connectors", "nested-connectors"] {
        let source = fs::read(fixtures.join(format!("deck-schema-v2-{name}.pptx")))?;
        let package = pptx_parse::parse_pptx_without_connectors(&source)?;
        assert!(!package.models_connectors());
        let session = DeckSession::from_package_with_source(package, &source, 273)?;
        {
            let mut txn = session.yrs_doc().transact_mut();
            let meta = txn.get_map("pptx:meta").unwrap();
            meta.insert(&mut txn, "schemaVersion", 2.0);
        }
        fs::write(
            fixtures.join(format!("deck-schema-v2-{name}.update.bin")),
            session.encode_state_as_update_v1(),
        )?;
        if name == "connectors" {
            session.move_shape(
                &EditCtx::local("fixture"),
                "slide:0:256",
                "slide:0:256:shape:1",
                952_500,
                1_047_750,
            )?;
            fs::write(
                fixtures.join("deck-schema-v2-connectors-moved.update.bin"),
                session.encode_state_as_update_v1(),
            )?;
            let current = DeckSession::open(&source, 277)?;
            fs::write(
                fixtures.join("deck-schema-v3-connectors.update.bin"),
                current.encode_state_as_update_v1(),
            )?;
            let migrated = DeckSession::open_from_update(
                &fs::read(fixtures.join("deck-schema-v2-connectors.update.bin"))?,
                278,
            )?;
            fs::write(
                fixtures.join("deck-schema-v3-legacy-connectors.update.bin"),
                migrated.encode_state_as_update_v1(),
            )?;
        }
    }
    let source =
        fs::read(branch.join("crates/pptx-parse/tests/fixtures/slide-number-fields.pptx"))?;
    let numbered = DeckSession::open(&source, 277)?;
    fs::write(
        fixtures.join("deck-schema-v3-slide-number-fields.update.bin"),
        numbered.encode_state_as_update_v1(),
    )?;
    println!("Generated legacy v2 and current v3 fixtures with current main");
    Ok(())
}
