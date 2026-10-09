use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, SplitReceipt};

pub fn seed(doc: &EditingDoc, kind: &str, split_ctx: &EditCtx) -> SplitReceipt {
    let plain = EditCtx::local("", "");
    doc.create_story("body", "oldtail", "Normal", "left")
        .unwrap();
    let split = doc
        .split_paragraph(split_ctx, Position::new("body", 3), None)
        .unwrap();
    match kind {
        "table" => {
            let table = doc
                .insert_table(&plain, Position::new("body", 4), 1, 1)
                .unwrap();
            doc.insert_text(
                &plain,
                Position::new(&table.created_story_ids[0], 0),
                "cell",
                FormatPolicy::Plain,
            )
            .unwrap();
        }
        "blockSdt" => {
            doc.create_story("control", "inside", "Normal", "left")
                .unwrap();
            doc.insert_embed(
                &plain,
                Position::new("body", 4),
                kind,
                vec![("story".to_owned(), "control".into())],
            )
            .unwrap();
        }
        _ => {
            doc.insert_embed(&plain, Position::new("body", 4), kind, vec![])
                .unwrap();
        }
    }
    split
}
