use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Loc, Position};
use docx_layout::types::LayoutBlock;

fn ctx() -> EditCtx {
    EditCtx::local("Owner", "2026-07-14T12:00:00Z")
}

fn texts(doc: &EditingDoc) -> Vec<String> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|para| para.text)
        .collect()
}

fn kinds(doc: &EditingDoc) -> Vec<&'static str> {
    yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default())
        .unwrap()
        .iter()
        .map(|block| match block {
            LayoutBlock::Paragraph(_) => "p",
            LayoutBlock::Table(_) => "table",
            _ => "other",
        })
        .collect()
}

#[test]
fn inline_content_at_the_start_of_a_paragraph_led_by_a_table_follows_the_table() {
    let inserts: [fn(&EditingDoc, Position); 3] = [
        |doc, at| {
            doc.insert_text(&ctx(), at, "x", FormatPolicy::Inherit)
                .unwrap();
        },
        |doc, at| {
            doc.insert_tab(&ctx(), at).unwrap();
        },
        |doc, at| {
            doc.insert_hard_break(&ctx(), at).unwrap();
        },
    ];
    for insert in inserts {
        let doc = EditingDoc::new(300);
        doc.create_story("body", "ab", "Normal", "left").unwrap();
        doc.insert_table(&ctx(), Position::new("body", 2), 1, 1)
            .unwrap();
        let para = doc.paragraphs("body").unwrap()[1].para_id.clone();
        insert(&doc, doc.locate(&Loc::new("body", para, 0)).unwrap());
        assert_eq!(kinds(&doc), ["p", "table", "p"]);
        assert_eq!(texts(&doc)[0], "ab");
    }
}
