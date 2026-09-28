use std::{env, fs, path::Path};

use pptx_edit::{DeckSession, EditCtx, TextStyle};
use yrs::{Map, Out, ReadTxn, Transact};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let fixtures = Path::new(&args[1]).join("crates/pptx-edit/tests/fixtures");
    for (name, client_id) in [("defaults", 2101), ("edits", 2102)] {
        let source = fs::read(fixtures.join(format!("deck-schema-v2.1-{name}.pptx")))?;
        let session = DeckSession::open(&source, client_id)?;
        {
            let txn = session.yrs_doc().transact();
            let meta = txn.get_map("pptx:meta").unwrap();
            assert_eq!(meta.get(&txn, "schemaVersion"), Some(Out::Any(2.1.into())));
        }
        if name == "edits" {
            let story = session.snapshot()?.slides[0].shapes[0].text_stories[0]
                .id
                .clone();
            let context = EditCtx::local("fixture");
            let typed = TextStyle {
                font_size_pt: Some(32.0),
                color: Some("#123456".to_owned()),
                ..TextStyle::default()
            };
            session.delete_text(&context, &story, 1, 3)?;
            session.delete_text(&context, &story, 9, 11)?;
            session.insert_text(&context, &story, 0, "NEW", &typed)?;
            session.delete_text(&context, &story, 27, 29)?;
            session.insert_text(&context, &story, 19, "\u{1D402}", &typed)?;
            let text: Vec<String> = session
                .story(&story)?
                .paragraphs
                .iter()
                .map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()).collect())
                .collect();
            assert_eq!(
                text,
                [
                    "NEWAha \u{1D401}eta gamma",
                    "\u{1D402}\u{394}elta \u{1F600} epsilon"
                ]
            );
        }
        fs::write(
            fixtures.join(format!("deck-schema-v2.1-{name}.update.bin")),
            session.encode_state_as_update_v1(),
        )?;
    }
    println!("Generated the released 2.1 fixtures");
    Ok(())
}
