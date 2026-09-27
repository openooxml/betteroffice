use std::{env, fs, path::Path};

use pptx_edit::DeckSession;
use yrs::{Map, Out, ReadTxn, Transact};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let fixtures = Path::new(&args[1]).join("crates/pptx-edit/tests/fixtures");
    let source = fs::read(fixtures.join("deck-schema-v2.1-defaults.pptx"))?;
    let session = DeckSession::open(&source, 2101)?;
    {
        let txn = session.yrs_doc().transact();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(meta.get(&txn, "schemaVersion"), Some(Out::Any(2.1.into())));
    }
    fs::write(
        fixtures.join("deck-schema-v2.1-defaults.update.bin"),
        session.encode_state_as_update_v1(),
    )?;
    println!("Generated the released 2.1 defaults fixture");
    Ok(())
}
